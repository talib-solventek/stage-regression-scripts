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
${NAT_PERIL_RATER_STATUS_XPATH}    //*[@id="ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlNatPerilRates_lblRaterStatus"]
${XOL_RATER_STATUS_XPATH}          //*[@id="ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlExcessOfLoss_lblRaterStatus"]


*** Tasks ***
Oil And Gas
    Open Workbook    ${DATA_SOURCE}
    ${excel_rows}=    Read Worksheet    OilAndGas
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
    Quote
    Account
    Verify Status
    ...    xpath=//span[contains(@id, 'labelStatus')]
    ...    Quoted
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
    # Close Browser

Login to App
    ${URL}=    Get From Dictionary    ${current_row}    A
    ${NUSER}=    Get From Dictionary    ${current_row}    B
    ${NUSERPASSWORD}=    Get From Dictionary    ${current_row}    C
    Open Browser    ${URL}    chrome    options=add_argument("--inprivate")
    Maximize Browser Window
    Set Selenium Implicit Wait    30s
    Press Keys    None    CTRL+R
    Safe Click Element    xpath=//a[normalize-space()='Create New Risk >']
    Safe Click Element    xpath=(//a[normalize-space()='Oil And Gas'])[1]
    Sleep    10s
    Safe Click Element
    ...    id:ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlClearanceSearch_CreateInsured
    Sleep    5s
    ${EMAIL}=    Set Variable    ${NUSER}@libertymutual.com
    Safe Wait Until Element Is Visible    xpath=//input[@type='email']    30s
    Clear Element Text    xpath=//input[@type='email']
    Input Text    xpath=//input[@type='email']    ${EMAIL}
    Sleep    1s
    Press Keys    xpath=//input[@type='email']    ENTER
    Safe Wait Until Element Is Visible    xpath=//input[@type='password']    30s
    Clear Element Text    xpath=//input[@type='password']
    Input Text    xpath=//input[@type='password']    ${NUSERPASSWORD}
    Sleep    1s
    Press Keys    xpath=//input[@type='password']    ENTER
    Sleep    5s

Risk Creation[First Run]
    Fill Insured Details[First Run]
    Verify Status
    ...    xpath=//span[contains(@id, 'labelStatus')]
    ...    Submission (SOV In Progress)
    Fill Pricing Details[First Run]

Fill Insured Details[First Run]
    Safe Click Element    (//span[normalize-space()='No'])[1]
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
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Safe Wait Until Element Is Visible
    ...    (//legend[@class='ui-widget ui-widget-header ui-corner-all'][normalize-space()='Additional Named Insured'])[1]
    ...    120s
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Safe Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]    120s
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListBranch'])[1]
    ...    Madrid
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListCompany'])[1]
    ...    SPAIN
    Select From List By Index
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListAssistant'])[1]
    ...    1
    ${department}=    Get From Dictionary    ${current_row}    G
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListDepartment'])[1]
    ...    ${department}
    ${producing_office}=    Get From Dictionary    ${current_row}    G
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListProducingOffice'])[1]
    ...    Madrid
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ddlPrimarySegmentOccupancy'])[1]
    ...    Ammonia
    ${effdate}=    Get From Dictionary    ${current_row}    H
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DatePickerEffectiveDate_textDate'])[1]
    ...    ${effdate}
    Safe Wait Until Element Is Visible    //*[@id="ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlRiskHeader_btnReturnToRiskSummary"]    15s
    ${risk_number}=    RPA.Browser.Selenium.Get Text    //*[@id="ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlRiskHeader_btnReturnToRiskSummary"]
    Evaluate    open(r'${CURDIR}${/}risk_numbers.txt', 'a').write('First Run - Risk Number: ${risk_number}\\n')
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Safe Wait Until Element Is Visible    (//legend[normalize-space()='Broker Info'])[1]    120s
    ${brokfirm}=    Get From Dictionary    ${current_row}    J
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirm_TextBoxBrokerFirm'])[1]
    ...    ${brokfirm}
    Safe Click Element
    ...    (//input[@id="ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirm_ButtonBrokerFirmSearchStartsWith"])[1]
    Sleep    10s
    Safe Click Element
    ...    xpath=(//table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']/tbody/tr/td//input[@type="submit"] | //table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']/tr/td//input[@type="submit"])[3]
    Execute Javascript    window.scrollTo(0, document.body.scrollHeight)
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Safe Wait Until Element Is Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_ctrlBrokerContactSearch_TextBoxBrokerContact'])[1]
    ...    120s
    ${brokcont}=    Get From Dictionary    ${current_row}    K
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_ctrlBrokerContactSearch_TextBoxBrokerContact'])[1]
    ...    ${brokcont}
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_ctrlBrokerContactSearch_ButtonBrokerContactSearchStartsWith'])[1]
    Sleep    20s
    Safe Click Element
    ...    xpath=(//table[@id='TableBrokerContactSearch']/tbody/tr/td//input[@value="Select"] | //table[@id='TableBrokerContactSearch']/tr/td//input[@value="Select"])[2]
    Safe Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_ButtonSave'])[1]
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]

Fill Pricing Details[First Run]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Location Management'])[1]    120s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlLocationManagement_btnImportLocationDialog'])[1]
    Sleep    10s
    Choose File
    ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlLocationManagement_fuLocationCSV'])[1]
    ...    ${CURDIR}${/}Location.xlsx
    Safe Click Element    (//span[normalize-space()='Import'])[1]
    Sleep    5s
    Safe Click Element    (//a[@class='btn btn-info'])[1]
    Wait Until Element Is Enabled    (//button[normalize-space()='Next'])[1]    120s
    Safe Click Element    (//button[normalize-space()='Next'])[1]
    Sleep    20s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlLocationManagement_btnAddLocation'])[1]
    Sleep    5s
    Safe Input Text    (//input[@id='Name'])[1]    test
    Safe Input Text    (//input[@id='country'])[1]    Bangladesh
    Safe Input Text    (//input[@id='Latitude'])[1]    78
    Safe Input Text    (//input[@id='Longitude'])[1]    170
    Safe Click Element    (//button[normalize-space()='Match & Cleanse'])[1]
    Sleep    5s
    Safe Click Element    (//button[normalize-space()='Next'])[1]
    Safe Click Element    (//a[normalize-space()='Location Details'])[1]
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlLocationInformation_ddlLocationType'])[1]
    ...    OGPC
    Sleep    2s
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlLocationInformation_ddlRatingSegment'])[1]
    ...    Chemicals
    Sleep    2s
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlLocationInformation_ddlOccupancy'])[1]
    ...    Air Separation / Oxygen
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlLocationInformation_ddlAssetCurrency'])[1]
    ...    USD
    ${pd_building}=    Get From Dictionary    ${current_row}    P
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlLocationInformation_txtPdBuilding'])[1]
    ...    500
    ${pd_mach_equip}=    Get From Dictionary    ${current_row}    P
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlLocationInformation_txtPdMachineryAndEquipment'])[1]
    ...    500
    ${pd_contents}=    Get From Dictionary    ${current_row}    P
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlLocationInformation_txtPdContents'])[1]
    ...    500
    ${pd_tiv}=    Get From Dictionary    ${current_row}    P
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlLocationInformation_txtPDTIV'])[1]
    ...    1500
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Yes'])[1]
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[2]
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[3]
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[4]
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_btnSave'])[1]
    Sleep    5s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl01'])[1]
    Sleep    60s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Ok'])[1]
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Operational Policy Terms'])[1]    120s
    ${lead_follow}=    Get From Dictionary    ${current_row}    Z
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlPolicyTerms_ddlLeadFollow'])[1]
    ...    ${lead_follow}
    ${policy_wording}=    Get From Dictionary    ${current_row}    AA
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlPolicyTerms_ddlPolicyWording'])[1]
    ...    ${policy_wording}
    ${combined_loss_limit}=    Get From Dictionary    ${current_row}    AB
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlPolicyTerms_txtCombinedLossLimit'])[1]
    ...    ${combined_loss_limit}
    ${MB_PD}=    Get From Dictionary    ${current_row}    AC
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlPolicyTerms_txtSublimitsMbPd'])[1]
    ...    ${MB_PD}
    ${MB_BI}=    Get From Dictionary    ${current_row}    AD
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlPolicyTerms_txtSublimitsMbBi'])[1]
    ...    ${MB_BI}
    ${cbi}=    Get From Dictionary    ${current_row}    AE
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlPolicyTerms_txtSublimitsCbi'])[1]
    ...    ${cbi}
    ${extra_expense}=    Get From Dictionary    ${current_row}    AF
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlPolicyTerms_txtSublimitsExtraExpense'])[1]
    ...    ${extra_expense}
    ${PD}=    Get From Dictionary    ${current_row}    AJ
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlPolicyTerms_txtDeductiblesPD'])[1]
    ...    ${PD}
    ${bi_days}=    Get From Dictionary    ${current_row}    AK
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlPolicyTerms_txtDeductiblesBIDays'])[1]
    ...    ${bi_days}
    ${mb_pd_ded}=    Get From Dictionary    ${current_row}    AL
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlPolicyTerms_txtMbPdDeductible'])[1]
    ...    ${mb_pd_ded}
    ${mb_bi_ded}=    Get From Dictionary    ${current_row}    AM
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlPolicyTerms_txtMbBiDeductibleDays'])[1]
    ...    ${mb_bi_ded}
    ${cbi_days}=    Get From Dictionary    ${current_row}    AN
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlPolicyTerms_txtDeductiblesCBIDays'])[1]
    ...    ${cbi_days}
    ${service_interuption}=    Get From Dictionary    ${current_row}    AO
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlPolicyTerms_ddlServiceInterruption'])[1]
    ...    ${service_interuption}
    ${bi_basis}=    Get From Dictionary    ${current_row}    AO
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlPolicyTerms_ddlBIBasis'])[1]
    ...    Fixed Costs
    ${bi_period}=    Get From Dictionary    ${current_row}    AO
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlPolicyTerms_ddlBIIndemnityPeriod'])[1]
    ...    5
    ${ip_start_date}=    Get From Dictionary    ${current_row}    AO
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlPolicyTerms_ddlIpStartDate'])[1]
    ...    Date of Loss
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Operational Location Terms'])[1]    120s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Operational Rates'])[1]    120s
    ${start_time}=    Get Time    epoch
    Sleep    20s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Natural Perils Policy Terms'])[1]    120s
    ${end_time}=    Get Time    epoch
    ${total_time}=    Subtract Time From Time    ${end_time}    ${start_time}
    Log    Total time taken on Operational Rates screen is ${total_time}
    ${sublimit_type1}=    Get From Dictionary    ${current_row}    AP
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlNaturalPerilTerms_rptNaturalPerilsTerms_ctl01_ddlSublimitType'])[1]
    ...    ${sublimit_type1}
    ${sublimit_type2}=    Get From Dictionary    ${current_row}    AQ
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlNaturalPerilTerms_rptNaturalPerilsTerms_ctl02_ddlSublimitType'])[1]
    ...    ${sublimit_type2}
    ${sublimit_type3}=    Get From Dictionary    ${current_row}    AR
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlNaturalPerilTerms_rptNaturalPerilsTerms_ctl03_ddlSublimitType'])[1]
    ...    ${sublimit_type3}
    ${bi_days1}=    Get From Dictionary    ${current_row}    AS
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlNaturalPerilTerms_rptNaturalPerilsTerms_ctl01_txtBiDays'])[1]
    ...    ${bi_days1}
    ${bi_days2}=    Get From Dictionary    ${current_row}    AT
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlNaturalPerilTerms_rptNaturalPerilsTerms_ctl02_txtBiDays'])[1]
    ...    ${bi_days2}
    ${bi_days3}=    Get From Dictionary    ${current_row}    AU
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlNaturalPerilTerms_rptNaturalPerilsTerms_ctl03_txtBiDays'])[1]
    ...    ${bi_days3}
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Natural Perils Location Terms'])[1]    120s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Natural Peril Rates'])[1]    120s
    Wait For Rater Status Complete    ${NAT_PERIL_RATER_STATUS_XPATH}
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Loss Experience'])[1]    120s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Discretionary Modifiers'])[1]    120s
    Safe Click Element
    ...    //*[@id="ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_divNavButtons"]/input[2]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Excess Of Loss'])[1]    120s
    Safe Click Element
    ...    //*[@id="ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlExcessOfLoss_btnXolRates"]
    Wait For Rater Status Complete    ${XOL_RATER_STATUS_XPATH}
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Summary Of Pricing'])[1]    120s
    ${liu_share}=    Get From Dictionary    ${current_row}    AV
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlSummaryOfPricing_txtLiuSharePercentage'])[1]
    ...    ${liu_share}
    ${commission}=    Get From Dictionary    ${current_row}    AW
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlSummaryOfPricing_txtCommission'])[1]
    ...    ${commission}
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    ${popup_visible}=    Run Keyword And Return Status    Safe Wait Until Element Is Visible    xpath=/html/body/div[2]    2s
    IF    ${popup_visible}
        Safe Click Element    xpath=/html/body/div[2]/div[3]/div/button/span
        Click Element    xpath=//body
        Sleep    5s
        Safe Click Element
        ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    END
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Fac'])[1]    120s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Terror'])[1]    120s
    ${other}=    Get From Dictionary    ${current_row}    AV
    Clear Element Text
    ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlTria_txtMarketCertified'])[1]
    ${market_certificate}=    Get From Dictionary    ${current_row}    AV
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlTria_txtMarketCertified'])[1]
    ...    1000
    ${garaet_prem}=    Get From Dictionary    ${current_row}    AX
    Sleep    5s
    Enter Terror All Other 100
    Click Element    xpath=//body
    Sleep    5s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='LibertyIndex'])[1]    120s

    Safe Click Element    (//span[normalize-space()='+ve'])[1]
    ${renewal_years}=    Get From Dictionary    ${current_row}    AX
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlLibertyIndex_txtRenewalYears'])[1]
    ...    ${renewal_years}
    ${claims_approach}=    Get From Dictionary    ${current_row}    AY
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlLibertyIndex_ddlClaimsApproach'])[1]
    ...    ${claims_approach}
    ${opinion}=    Get From Dictionary    ${current_row}    AZ
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlLibertyIndex_ddlExecOpinion'])[1]
    ...    ${opinion}
    ${cyber_type}=    Get From Dictionary    ${current_row}    BA
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlLibertyIndex_ddlCyberType'])[1]
    ...    ${cyber_type}
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Subjectivities'])[1]    120s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]

Fill Insured Details[Copy Risk]
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    Sleep    5s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Safe Wait Until Element Is Visible
    ...    (//legend[@class='ui-widget ui-widget-header ui-corner-all'][normalize-space()='Additional Named Insured'])[1]
    ...    120s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Safe Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]    120s
    Select From List By Index
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListAssistant'])[1]
    ...    1
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ddlPrimarySegmentOccupancy'])[1]
    ...    Ammonia
    Safe Wait Until Element Is Visible    //*[@id="ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlRiskHeader_btnReturnToRiskSummary"]    15s
    ${risk_number}=    RPA.Browser.Selenium.Get Text    //*[@id="ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlRiskHeader_btnReturnToRiskSummary"]
    Evaluate    open(r'${CURDIR}${/}risk_numbers.txt', 'a').write('Copy Risk - Risk Number: ${risk_number}\\n')
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Safe Wait Until Element Is Visible    (//legend[normalize-space()='Broker Info'])[1]    120s
    ${brokfirm}=    Get From Dictionary    ${current_row}    J
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirm_TextBoxBrokerFirm'])[1]
    ...    ${brokfirm}
    Safe Click Element
    ...    (//input[@id="ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirm_ButtonBrokerFirmSearchStartsWith"])[1]
    Sleep    10s
    Safe Click Element
    ...    xpath=(//table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']/tbody/tr/td//input[@type="submit"] | //table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']/tr/td//input[@type="submit"])[3]
    Execute Javascript    window.scrollTo(0, document.body.scrollHeight)
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Safe Wait Until Element Is Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_ctrlBrokerContactSearch_TextBoxBrokerContact'])[1]
    ...    120s
    ${brokcont}=    Get From Dictionary    ${current_row}    K
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_ctrlBrokerContactSearch_TextBoxBrokerContact'])[1]
    ...    ${brokcont}
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_ctrlBrokerContactSearch_ButtonBrokerContactSearchStartsWith'])[1]
    Sleep    10s
    Safe Click Element
    ...    xpath=(//table[@id='TableBrokerContactSearch']/tbody/tr/td//input[@value="Select"] | //table[@id='TableBrokerContactSearch']/tr/td//input[@value="Select"])[2]
    Safe Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_ButtonSave'])[1]
    Safe Click Element    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]

Fill Pricing Details[Copy Risk]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Location Management'])[1]    120s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Sleep    5s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Ok'])[1]
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Operational Policy Terms'])[1]    120s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Operational Location Terms'])[1]    120s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Operational Rates'])[1]    120s
    Sleep    20s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Natural Perils Policy Terms'])[1]
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Natural Perils Location Terms'])[1]    120s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Natural Peril Rates'])[1]    120s
    Wait For Rater Status Complete    ${NAT_PERIL_RATER_STATUS_XPATH}
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Loss Experience'])[1]    120s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Discretionary Modifiers'])[1]    120s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Excess Of Loss'])[1]    120s
    Safe Click Element
    ...    //*[@id="ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlExcessOfLoss_btnXolRates"]
    Wait For Rater Status Complete    ${XOL_RATER_STATUS_XPATH}
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Summary Of Pricing'])[1]    120s
    ${liu_share}=    Get From Dictionary    ${current_row}    AV
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlSummaryOfPricing_txtLiuSharePercentage'])[1]
    ...    ${liu_share}
    ${commission}=    Get From Dictionary    ${current_row}    AW
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlSummaryOfPricing_txtCommission'])[1]
    ...    ${commission}
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    ${popup_visible}=    Run Keyword And Return Status    Safe Wait Until Element Is Visible    xpath=/html/body/div[2]    2s
    IF    ${popup_visible}
        Safe Click Element    xpath=/html/body/div[2]/div[3]/div/button/span
        Click Element    xpath=//body
        Sleep    5s
        Safe Click Element
        ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    END
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Fac'])[1]    120s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Terror'])[1]    120s
    Enter Terror All Other 100
    Sleep    5s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='LibertyIndex'])[1]    120s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Subjectivities'])[1]    120s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]

Risk Creation[Renewal]
    Fill Insured Details[Renewal]
    Verify Status
    ...    xpath=//span[contains(@id, 'labelStatus')][1]
    ...    Submission
    Fill Pricing Details[Renewal]

Fill Insured Details[Renewal]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Insured Details'])[1]    120s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Safe Wait Until Element Is Visible
    ...    (//legend[@class='ui-widget ui-widget-header ui-corner-all'][normalize-space()='Additional Named Insured'])[1]
    ...    120s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Safe Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]    120s
    Select From List By Index
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ddlPrimarySegmentOccupancy'])[1]
    ...    1
    Safe Wait Until Element Is Visible    //*[@id="ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlRiskHeader_btnReturnToRiskSummary"]    15s
    ${risk_number}=    RPA.Browser.Selenium.Get Text    //*[@id="ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlRiskHeader_btnReturnToRiskSummary"]
    Evaluate    open(r'${CURDIR}${/}risk_numbers.txt', 'a').write('Renewal - Risk Number: ${risk_number}\\n')
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]

Fill Pricing Details[Renewal]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Location Management'])[1]    120s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Ok'])[1]
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Operational Policy Terms'])[1]    120s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Operational Location Terms'])[1]    120s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Operational Rates'])[1]    120s
    Sleep    20s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Natural Perils Policy Terms'])[1]
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Natural Perils Location Terms'])[1]    120s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Natural Peril Rates'])[1]    120s
    Wait For Rater Status Complete    ${NAT_PERIL_RATER_STATUS_XPATH}
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Loss Experience'])[1]    120s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Discretionary Modifiers'])[1]    120s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Excess Of Loss'])[1]    120s
    Safe Click Element
    ...    //*[@id="ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlExcessOfLoss_btnXolRates"]
    Wait For Rater Status Complete    ${XOL_RATER_STATUS_XPATH}
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Summary Of Pricing'])[1]    120s
    ${liu_share}=    Get From Dictionary    ${current_row}    AV
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlSummaryOfPricing_txtLiuSharePercentage'])[1]
    ...    ${liu_share}
    ${commission}=    Get From Dictionary    ${current_row}    AW
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlSummaryOfPricing_txtCommission'])[1]
    ...    ${commission}
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    ${popup_visible}=    Run Keyword And Return Status    Safe Wait Until Element Is Visible    xpath=/html/body/div[2]    2s
    IF    ${popup_visible}
        Safe Click Element    xpath=/html/body/div[2]/div[3]/div/button/span
        Click Element    xpath=//body
        Sleep    5s
        Safe Click Element
        ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    END
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Fac'])[1]    120s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Terror'])[1]    120s
    ${is_market_cert_visible}=    Is Element Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlTria_txtMarketCertified'])[1]
    IF    ${is_market_cert_visible}
        Clear Element Text
        ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlTria_txtMarketCertified'])[1]
        ${market_certificate}=    Get From Dictionary    ${current_row}    AV
        Safe Input Text
        ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlTria_txtMarketCertified'])[1]
        ...    1000
    END
    ${garaet_prem}=    Get From Dictionary    ${current_row}    AX
    ${is_gareat_visible}=    Is Element Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlFrenchTerrorAndCatNat_txtGareatPremiumAmount'])[1]
    IF    ${is_gareat_visible}
        Safe Input Text
        ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlFrenchTerrorAndCatNat_txtGareatPremiumAmount'])[1]
        ...    10000
    END
    Sleep    5s
    Enter Terror All Other 100
    Sleep    5s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='LibertyIndex'])[1]    120s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Subjectivities'])[1]    120s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]

Fill Pricing Details[Reissue]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Location Management'])[1]    120s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Sleep    5s
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Operational Policy Terms'])[1]    120s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Operational Location Terms'])[1]    120s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Operational Rates'])[1]    120s
    Sleep    20s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Natural Perils Policy Terms'])[1]
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Natural Perils Location Terms'])[1]    120s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Natural Peril Rates'])[1]    120s
    Wait For Rater Status Complete    ${NAT_PERIL_RATER_STATUS_XPATH}
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Loss Experience'])[1]    120s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Discretionary Modifiers'])[1]    120s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Excess Of Loss'])[1]    120s
    Execute Javascript    window.scrollTo(0, document.body.scrollHeight)
    Wait For Rater Status Complete    ${XOL_RATER_STATUS_XPATH}
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Summary Of Pricing'])[1]    120s
    ${liu_share}=    Get From Dictionary    ${current_row}    AV
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlSummaryOfPricing_txtLiuSharePercentage'])[1]
    ...    ${liu_share}
    ${commission}=    Get From Dictionary    ${current_row}    AW
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlSummaryOfPricing_txtCommission'])[1]
    ...    ${commission}
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Fac'])[1]    120s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Terror'])[1]    120s
    Enter Terror All Other 100
    Sleep    5s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='LibertyIndex'])[1]    120s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Subjectivities'])[1]    120s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]

Pre Bind Endorsement
    Safe Click Element    (//input[@id='btnShowAllEndorsementTypes'])[1]
    ${endorsement_type}=    Get From Dictionary    ${current_row}    AP
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlPreBindEndorsements_txtEndorsementType'])[1]
    ...    ${endorsement_type}
    Sleep    2s
    Safe Click Element    (//a[normalize-space()='${endorsement_type}'])[1]
    Sleep    2s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlPreBindEndorsements_btAddEndorsement'])[1]
    Sleep    3s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Subjectivities'])[1]    2s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]
    Sleep    5s

Quote
    Sleep    5s
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Quote Final'])[1]    120s
    Sleep    5s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_btnComplete'])[1]

Quote[Copy Risk]
    Sleep    5s
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Quote Final'])[1]    120s
    Sleep    5s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_btnGenerateQuote'])[1]

Pre Bind Endorsement[Reissue]
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Sleep    5s
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Subjectivities'])[1]    2s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Sleep    5s

Quote[Reissue]
    Sleep    5s
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Quote Final'])[1]    120s
    Sleep    5s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_btnGenerateQuote'])[1]

Quote[Renewal]
    Sleep    5s
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Quote Final'])[1]    120s
    Sleep    5s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_btnComplete'])[1]

Ready To Bind
    Safe Click Element    (//a[normalize-space()='Bind'])[1]
    Safe Wait Until Element Is Visible    (//legend[@id='fsMainContentLegend'])[1]    30s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl01'])[1]
    Safe Wait Until Element Is Visible    (//legend[@id='fsMainContentLegend'])[1]    30s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_btnNext'])[1]
    ${written_line}=    Get From Dictionary    ${current_row}    CC
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlBind_txtWrittenLinePercent'])[1]
    ...    ${written_line}
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[2]
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_btnBind'])[1]

Book
    [Arguments]    ${download_policy}
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Book and Issue'])[1]    30s
    Safe Click Element    (//a[normalize-space()='Book and Issue'])[1]
    Safe Wait Until Element Is Visible    (//legend[@id='fsMainContentLegend'])[1]    30s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_btnNext'])[1]
    Safe Wait Until Element Is Visible    (//legend[@id='fsMainContentLegend'])[1]    30s
    ${has_ci}=    Evaluate    'CI' in $current_row
    IF    ${has_ci}
        ${prim_occupancy}=    Get From Dictionary    ${current_row}    CI
    ELSE
        ${prim_occupancy}=    Set Variable    Ammonia
    END
    Run Keyword And Ignore Error    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_ctrlPremiumBreakdownContainer_ddlPrimarySegmentOccupancy'])[1]
    ...    ${prim_occupancy}
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ContentPlaceHolderMain_btnNext'])[1]
    Safe Wait Until Element Is Visible    (//legend[@id='fsMainContentLegend'])[1]    30s
    Sleep    60s
    ${is_umr_visible}=    Is Element Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlBook_TextBoxUmr' or contains(@id, 'ctrlBook_TextBoxUmr')])[1]
    IF    ${is_umr_visible}
        ${umr}=    Get From Dictionary    ${current_row}    AZ
        Safe Input Text
        ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlBook_TextBoxUmr' or contains(@id, 'ctrlBook_TextBoxUmr')])[1]
        ...    ${umr}
    END
    ${is_no_visible}=    Is Element Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    IF    ${is_no_visible}
        Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
        Sleep    5s
    END
    ${is_due_days_visible}=    Is Element Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlBook_ctrlBookInstallmentPlan_tbDueDays' or contains(@id, 'ctrlBook_ctrlBookInstallmentPlan_tbDueDays')])[1]
    IF    ${is_due_days_visible}
        ${due_days}=    Get From Dictionary    ${current_row}    AZ
        Safe Input Text
        ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlBook_ctrlBookInstallmentPlan_tbDueDays' or contains(@id, 'ctrlBook_ctrlBookInstallmentPlan_tbDueDays')])[1]
        ...    ${due_days}
    END
    ${is_next_visible}=    Is Element Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ButtonNext' or contains(@id, 'ButtonNext')])[1]
    IF    ${is_next_visible}
        Safe Click Element
        ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ButtonNext' or contains(@id, 'ButtonNext')])[1]
        Safe Wait Until Element Is Visible    (//legend[@id='fsMainContentLegend'])[1]    30s
        ${is_yes1}=    Is Element Visible    (//span[@class='ui-button-text'][normalize-space()='Yes'])[1]
        IF    ${is_yes1}
            Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Yes'])[1]
        END
        ${is_yes2}=    Is Element Visible    (//span[@class='ui-button-text'][normalize-space()='Yes'])[2]
        IF    ${is_yes2}
            Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Yes'])[2]
        END
        ${is_yes3}=    Is Element Visible    (//span[@class='ui-button-text'][normalize-space()='Yes'])[3]
        IF    ${is_yes3}
            Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Yes'])[3]
        END
    END
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ButtonBook' or contains(@id, 'ButtonBook')])[1]

Copy Quote
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Copy Quote'])[1]    30s
    Safe Click Element    (//a[normalize-space()='Copy Quote'])[1]
    Safe Wait Until Element Is Visible
    ...    xpath=(//span[contains(@id, 'HeaderQuote')])[2]
    ...    10s
    Verify Status
    ...    xpath=(//span[contains(@id, 'HeaderQuote')])[2]
    ...    Option 2 - [Status : Quoted (In Revision)]
    Safe Click Element
    ...    xpath=(//span[contains(@id, 'HeaderQuote')])[1]
    Sleep    5s

View Risk
    Safe Wait Until Element Is Visible    (//*[normalize-space()='View Risk'])[1]    30s
    Safe Click Element    (//a[normalize-space()='View Risk'])[1]
    Sleep    5s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonViewCurrent'])[1]
    Verify Status    (//legend[@id='fsMainContentLegend'])[1]    View Risk

Post Bind Endorsement
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Endorsement'])[1]    30s
    Safe Click Element    (//a[normalize-space()='Endorsement'])[1]
    ${endorsement_type_list}=    Get From Dictionary    ${current_row}    CE
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_txtEndorsementType'])[1]
    ...    ${endorsement_type_list}
    Sleep    2s
    Safe Click Element    (//a[normalize-space()='${endorsement_type_list}'])[1]
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_btAddEndorsement'])[1]
    Sleep    3s
    ${endorsement_effective_date}=    Get From Dictionary    ${current_row}    CF
    ${endorsement_effective_date}=    Evaluate
    ...    $endorsement_effective_date.strftime('%d/%m/%Y') if hasattr($endorsement_effective_date, 'strftime') else __import__('datetime').datetime.strptime(str($endorsement_effective_date).split('.')[0].strip(), '%Y-%m-%d %H:%M:%S' if ' ' in str($endorsement_effective_date) else '%Y-%m-%d').strftime('%d/%m/%Y') if '-' in str($endorsement_effective_date) else str($endorsement_effective_date)
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderEndorsementMain_ContentPlaceHolderEndorsement_DatePickerEndorsementEffectiveDate_textDate'])[1]
    ...    ${endorsement_effective_date}
    TRY
        ${expiry_date}=    Get Current Date    increment=5 days    result_format=%d/%m/%Y
        Safe Input Text
        ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderEndorsementMain_ContentPlaceHolderEndorsement_DatePickerPolicyExpiryDate_textDate'])[1]
        ...    ${expiry_date}
    EXCEPT
        ${expiry_date}=    Get Current Date    increment=5 days    result_format=%d/%m/%Y
        Safe Input Text
        ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderEndorsementMain_ContentPlaceHolderEndorsement_DatePickerPolicyExpiryDate_textDate'])[1]
        ...    ${expiry_date}
    END
    ${due_days}=    Get From Dictionary    ${current_row}    CH
    Safe Input Text
    ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderEndorsementMain_ContentPlaceHolderEndorsement_tbDueDays'])[1]
    ...    ${due_days}
    Sleep    2s
    Click Element    xpath=//body
    Run Keyword And Ignore Error    Execute Javascript    if (document.activeElement) { document.activeElement.blur(); }
    Sleep    3s
    Safe Wait Until Element Is Visible
    ...    xpath=//*[@id="ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_endorsementStandardButtons_btnSubmit"]
    ...    30s
    Safe Click Element
    ...    xpath=//*[@id="ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_endorsementStandardButtons_btnSubmit"]
    Sleep    5s
    Safe Wait Until Element Is Visible
    ...    xpath=//*[@id="ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_BtnBack"]
    ...    60s
    Safe Click Element
    ...    xpath=//*[@id="ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_BtnBack"]
    Sleep    5s

Renewal
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Renew'])[1]    30s
    Safe Click Element    (//a[normalize-space()='Renew'])[1]
    Verify Status
    ...    xpath=//span[contains(@id, 'labelStatus')]
    ...    Submission
    Risk Creation[Renewal]
    Quote[Renewal]
    Account[Renewal]
    Ready To Bind
    Book    False

Reissue
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Reissue'])[1]    30s
    Safe Click Element    (//a[normalize-space()='Reissue'])[1]
    Sleep    10s
    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Yes'])[7]
    Verify Status
    ...    xpath=//span[contains(@id, 'labelStatus')]
    ...    Quoted (Pending, In Revision, RI)
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Edit Quote'])[1]    30s
    Safe Click Element    (//a[normalize-space()='Edit Quote'])[1]
    Fill Pricing Details[Reissue]
    Quote[Reissue]
    Ready To Bind
    Book    False

Copy Risk
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Copy Risk'])[1]    30s
    Safe Click Element    (//a[normalize-space()='Copy Risk'])[1]
    Sleep    10s
    Safe Click Element    //span[contains(.,"Original Risks's Account")]
    Fill Insured Details[Copy Risk]
    Fill Pricing Details[Copy Risk]
    Quote[Copy Risk]
    Ready To Bind
    Book    False

Account
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Account'])[1]    30s
    Safe Click Element    (//a[normalize-space()='Account'])[1]
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlRiskAppraisal_ctrlRiskAppraisalButtons_btnNewRiskAppraisalRequest'])[1]
    ${risk_engineer}=    Get From Dictionary    ${current_row}    BG
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlRiskAppraisal_ctrlRiskAppraisalButtons_ddlRiskEngineer'])[1]
    ...    ${risk_engineer}
    ${Comments}=    Get From Dictionary    ${current_row}    BH
    Safe Input Text
    ...    (//textarea[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlRiskAppraisal_ctrlRiskAppraisalButtons_txtDescription'])[1]
    ...    ${Comments}
    Safe Click Element    (//span[normalize-space()='Create'])[1]
    ${account_number}=    RPA.Browser.Selenium.Get Text
    ...    (//span[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlAccountHeader_labelAccountNumber'])[1]
    Set Global Variable
    ...    ${url_ra}
    ...    https://test-riskengineering.grsazr.lmig.com/RiskAppraisal/${account_number}/Edit
    RA Approval    ${account_number}
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnReturnToRiskSummary'])[1]
    Safe Click Element    (//a[normalize-space()='Edit Submission'])[1]
    Safe Click Element    (//a[normalize-space()='Pricing'])[1]
    Sleep    5s
    Safe Click Element    (//a[normalize-space()='Operational'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Operational Policy Terms'])[1]    120s
    Safe Click Element    (//span[normalize-space()='Full'])[1]
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Operational Location Terms'])[1]    120s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Operational Rates'])[1]    120s
    Sleep    20s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Natural Perils Policy Terms'])[1]
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Natural Perils Location Terms'])[1]    120s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Natural Peril Rates'])[1]    120s
    Wait For Rater Status Complete    ${NAT_PERIL_RATER_STATUS_XPATH}
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Loss Experience'])[1]    120s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Discretionary Modifiers'])[1]    120s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Excess Of Loss'])[1]    120s
    Execute Javascript    window.scrollTo(0, document.body.scrollHeight)
    Wait For Rater Status Complete    ${XOL_RATER_STATUS_XPATH}
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Summary Of Pricing'])[1]    120s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Fac'])[1]    120s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Terror'])[1]    120s
    Enter Terror All Other 100
    Sleep    5s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='LibertyIndex'])[1]    120s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Subjectivities'])[1]    120s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]
    Sleep    5s
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Quote Final'])[1]    120s
    Sleep    5s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_btnGenerateQuote'])[1]
    Sleep    30s

Account[Renewal]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Account'])[1]    30s
    Safe Click Element    (//a[normalize-space()='Account'])[1]
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlRiskAppraisal_ctrlRiskAppraisalButtons_btnNewRiskAppraisalRequest'])[1]
    ${risk_engineer}=    Get From Dictionary    ${current_row}    BG
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlRiskAppraisal_ctrlRiskAppraisalButtons_ddlRiskEngineer'])[1]
    ...    ${risk_engineer}
    ${Comments}=    Get From Dictionary    ${current_row}    BH
    Safe Input Text
    ...    (//textarea[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlRiskAppraisal_ctrlRiskAppraisalButtons_txtDescription'])[1]
    ...    ${Comments}
    Safe Click Element    (//span[normalize-space()='Create'])[1]
    ${account_number}=    RPA.Browser.Selenium.Get Text
    ...    (//span[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlAccountHeader_labelAccountNumber'])[1]
    Set Global Variable
    ...    ${url_ra}
    ...    https://test-riskengineering.grsazr.lmig.com/RiskAppraisal/${account_number}/Edit
    RA Approval[Renewal]    ${account_number}
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnReturnToRiskSummary'])[1]
    Safe Click Element    (//a[normalize-space()='Edit Submission'])[1]
    Safe Click Element    (//a[normalize-space()='Pricing'])[1]
    Sleep    5s
    Safe Click Element    (//a[normalize-space()='Operational'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Operational Policy Terms'])[1]    120s
    Safe Click Element    (//span[normalize-space()='Full'])[1]
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Operational Location Terms'])[1]    120s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Operational Rates'])[1]    120s
    Sleep    20s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Natural Perils Policy Terms'])[1]
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Natural Perils Location Terms'])[1]    120s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Natural Peril Rates'])[1]    120s
    Wait For Rater Status Complete    ${NAT_PERIL_RATER_STATUS_XPATH}
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Loss Experience'])[1]    120s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Discretionary Modifiers'])[1]    120s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Excess Of Loss'])[1]    120s
    Execute Javascript    window.scrollTo(0, document.body.scrollHeight)
    Wait For Rater Status Complete    ${XOL_RATER_STATUS_XPATH}
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Summary Of Pricing'])[1]    120s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Fac'])[1]    120s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Terror'])[1]    120s
    Enter Terror All Other 100
    Sleep    5s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='LibertyIndex'])[1]    120s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ContentPlaceHolderMain$ctl00'])[1]
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Subjectivities'])[1]    120s
    Safe Click Element
    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]
    Sleep    5s
    Safe Wait Until Element Is Visible    (//*[normalize-space()='Quote Final'])[1]    120s
    Sleep    5s
    Safe Click Element
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_btnGenerateQuote'])[1]
    Sleep    30s

Continue on Next Risk
    Safe Wait Until Element Is Visible    xpath=//a[normalize-space()='Create New Risk >']    120s
    Safe Click Element    xpath=//a[normalize-space()='Create New Risk >']
    Safe Click Element    xpath=//a[normalize-space()='Create New Risk >']
    Safe Wait Until Element Is Visible    xpath=//a[normalize-space()='US Primary GL']    30s
    Safe Click Element    xpath=//a[normalize-space()='US Primary GL']
    Safe Click Element
    ...    id:ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlClearanceSearch_CreateInsured
    Safe Wait Until Element Is Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[1]    40s

Verify Status
    [Arguments]    ${location}    ${status_expected}
    Wait Until Keyword Succeeds    45x    2s    Element Should Contain    ${location}    ${status_expected}

RA Approval
    [Arguments]    ${account_number}
    ${URL_RA1}=    Get From Dictionary    ${current_row}    BI
    ${NUSER}=    Get From Dictionary    ${current_row}    B
    ${NUSERPASSWORD}=    Get From Dictionary    ${current_row}    C
    Open Browser    ${URL_RA1}    chrome    options=add_argument("--inprivate")
    Sleep    10s
    Safe Click Element    (//a[@class='btn btn-info'])[1]
    Safe Input Text    (//input[@id='i0116'])[1]    n9970632@libertymutual.com
    Safe Click Element    (//input[@id='idSIButton9'])[1]
    Safe Input Text    (//input[@id='i0118'])[1]    wm$qY@32
    Safe Click Element    (//input[@id='idSIButton9'])[1]
    Sleep    10s
    ${locator}=    Send Keys    keys=${NUSER}{TAB}${NUSERPASSWORD}    send_enter=${TRUE}
    Send Keys    keys={RETURN}    send_enter=${TRUE}
    Sleep    20s
    Safe Click Element    //*[@id="navbar-collapse-1"]/ul/li[1]/a
    Sleep    5s
    Safe Click Element    (//span[@class='k-icon k-i-filter'])[8]
    Safe Input Text    (//input[@title='Value'])[1]    ${account_number}
    Safe Click Element    (//button[normalize-space()='Filter'])[1]
    Sleep    10s
    Safe Click Element    (//a[normalize-space()='Create RA'])[1]
    Sleep    20s
    Safe Click Element    (//td[normalize-space()='test'])[1]
    Sleep    5s
    Safe Click Element    (//a[normalize-space()='Executive Summary'])[1]
    Sleep    10s
    FOR    ${i}    IN RANGE    10
        Press Keys    None    ARROW_DOWN
    END
    Select Frame    xpath=//iframe[@title='Rich Text Editor, BusinessDescription']
    Safe Wait Until Element Is Visible    xpath=//body
    Click Element    xpath=//body
    Press Keys    xpath=//body    test
    Unselect Frame
    Sleep    8s

    Select Frame    xpath=//iframe[@title='Rich Text Editor, ExecutiveSummary']
    Safe Wait Until Element Is Visible    xpath=//body
    Click Element    xpath=//body
    Press Keys    xpath=//body    test
    Unselect Frame
    Sleep    5s
    
    Safe Click Element    (//a[normalize-space()='RA Summary'])[1]
    Safe Click Element    //span[@class='glyphicon glyphicon-edit']
    Safe Click Element    (//a[normalize-space()='Loss Estimates'])[1]
    Safe Click Element    (//a[normalize-space()='EML'])[1]
    Safe Input Text    (//input[@id='PropertyDamageOGPCFields_PropertyDamageValue'])[1]    10000
    Safe Input Text
    ...    (//input[@id='PropertyDamageOGPCFields_AdditionalDamagePercentage'])[1]
    ...    100
    Safe Click Element    (//a[normalize-space()='PML'])[1]
    Sleep    5s
    Safe Input Text    (//input[@id='PropertyDamageOGPCFields_PropertyDamageValue'])[1]    10000
    Safe Input Text
    ...    (//input[@id='PropertyDamageOGPCFields_AdditionalDamagePercentage'])[1]
    ...    100
    Safe Click Element    (//a[normalize-space()='RA Score'])[1]
    Safe Click Element    (//input[@id='convertQuick'])[1]
    Sleep    5s
    Safe Input Text    (//input[@id='Groups_0__Categories_1__QuickScore'])[1]    90
    Safe Input Text    (//input[@id='Groups_1__Categories_0__QuickScore'])[1]    89
    Safe Input Text    (//input[@id='Groups_1__Categories_1__QuickScore'])[1]    89
    Safe Input Text    (//input[@id='Groups_1__Categories_2__QuickScore'])[1]    89
    Safe Input Text    (//input[@id='Groups_1__Categories_3__QuickScore'])[1]    89
    Safe Input Text    (//input[@id='Groups_1__Categories_4__QuickScore'])[1]    89
    Safe Input Text    (//input[@id='Groups_1__Categories_5__QuickScore'])[1]    89
    Sleep    10s
    Safe Input Text    (//input[@id='Groups_2__Categories_0__QuickScore'])[1]    3
    Safe Click Element    (//a[normalize-space()='Location Summary'])[1]
    Sleep    5s
    Safe Click Element    (//button[normalize-space()='Complete Location Risk Appraisal'])[1]
    Safe Click Element    (//a[normalize-space()='Natural Perils'])[1]
    Safe Click Element    (//input[@id='ReviewedByRiskEngineer'])[1]
    Safe Click Element    (//button[normalize-space()='Complete'])[1]
    Sleep    5s
    Get Browser Ids
    Switch Browser    1

RA Approval[Renewal]
    [Arguments]    ${account_number}
    ${URL_RA1}=    Get From Dictionary    ${current_row}    BI
    ${NUSER}=    Get From Dictionary    ${current_row}    B
    ${NUSERPASSWORD}=    Get From Dictionary    ${current_row}    C
    Open Browser    ${URL_RA1}    chrome    options=add_argument("--inprivate")
    Sleep    10s
    Get Browser Ids
    Switch Browser    3
    Safe Click Element    (//a[@class='btn btn-info'])[1]
    Safe Input Text    (//input[@id='i0116'])[1]    n9970632@libertymutual.com
    Safe Click Element    (//input[@id='idSIButton9'])[1]
    Safe Input Text    (//input[@id='i0118'])[1]    wm$qY@32
    Safe Click Element    (//input[@id='idSIButton9'])[1]
    Sleep    10s
    Switch Browser    3
    ${locator}=    Send Keys    keys=${NUSER}{TAB}${NUSERPASSWORD}    send_enter=${TRUE}
    Send Keys    keys={RETURN}    send_enter=${TRUE}
    Sleep    120s
    Safe Click Element    //*[@id="navbar-collapse-1"]/ul/li[1]/a
    Sleep    5s
    Safe Click Element    (//span[@class='k-icon k-i-filter'])[8]
    Safe Input Text    (//input[@title='Value'])[1]    ${account_number}
    Safe Click Element    (//button[normalize-space()='Filter'])[1]
    Sleep    10s
    Safe Click Element    (//a[normalize-space()='Renew RA'])[1]
    Sleep    5s
    Safe Click Element    (//a[normalize-space()='Executive Summary'])[1]
    Sleep    20s
    Safe Click Element    (//a[normalize-space()='RA Summary'])[1]
    Safe Click Element    //span[@class='glyphicon glyphicon-edit']
    Safe Click Element    (//a[normalize-space()='Loss Estimates'])[1]
    Safe Click Element    (//a[normalize-space()='EML'])[1]
    Safe Input Text    (//input[@id='PropertyDamageOGPCFields_PropertyDamageValue'])[1]    10000
    Safe Input Text
    ...    (//input[@id='PropertyDamageOGPCFields_AdditionalDamagePercentage'])[1]
    ...    100
    Safe Click Element    (//a[normalize-space()='PML'])[1]
    Sleep    5s
    Safe Input Text    (//input[@id='PropertyDamageOGPCFields_PropertyDamageValue'])[1]    10000
    Safe Input Text
    ...    (//input[@id='PropertyDamageOGPCFields_AdditionalDamagePercentage'])[1]
    ...    100
    Safe Click Element    (//a[normalize-space()='RA Score'])[1]
    Safe Click Element    (//input[@id='convertQuick'])[1]
    Sleep    5s
    Safe Input Text    (//input[@id='Groups_0__Categories_1__QuickScore'])[1]    90
    Safe Input Text    (//input[@id='Groups_1__Categories_0__QuickScore'])[1]    89
    Safe Input Text    (//input[@id='Groups_1__Categories_1__QuickScore'])[1]    89
    Safe Input Text    (//input[@id='Groups_1__Categories_2__QuickScore'])[1]    89
    Safe Input Text    (//input[@id='Groups_1__Categories_3__QuickScore'])[1]    89
    Safe Input Text    (//input[@id='Groups_1__Categories_4__QuickScore'])[1]    89
    Safe Input Text    (//input[@id='Groups_1__Categories_5__QuickScore'])[1]    89
    Sleep    10s
    Safe Input Text    (//input[@id='Groups_2__Categories_0__QuickScore'])[1]    3
    Safe Click Element    (//a[normalize-space()='Location Summary'])[1]
    Sleep    5s
    Safe Click Element    (//button[normalize-space()='Complete Location Risk Appraisal'])[1]
    Safe Click Element    (//a[normalize-space()='Natural Perils'])[1]
    Safe Click Element    (//input[@id='ReviewedByRiskEngineer'])[1]
    Safe Click Element    (//button[normalize-space()='Complete'])[1]
    Sleep    5s
    Get Browser Ids
    Switch Browser    1

Safe Click Element
    [Arguments]    ${locator}    ${retries}=15x    ${retry_interval}=2s
    Wait Until Keyword Succeeds
    ...    ${retries}
    ...    ${retry_interval}
    ...    RPA.Browser.Selenium.Click Element When Visible
    ...    ${locator}

Safe Input Text
    [Arguments]    ${locator}    ${text}    ${retries}=15x    ${retry_interval}=2s
    Wait Until Keyword Succeeds
    ...    ${retries}
    ...    ${retry_interval}
    ...    RPA.Browser.Selenium.Input Text When Element Is Visible
    ...    ${locator}    ${text}

Check Element Visibility Safely
    [Arguments]    ${locator}
    ${status}    ${is_visible}=    Run Keyword And Ignore Error    RPA.Browser.Selenium.Is Element Visible    ${locator}
    IF    '${status}' == 'FAIL'
        Fail    Element not visible yet
    END
    IF    not $is_visible
        Fail    Element not visible yet
    END

Safe Wait Until Element Is Visible
    [Arguments]    ${locator}    ${timeout}=120s    ${retry_interval}=2s
    Wait Until Keyword Succeeds
    ...    ${timeout}
    ...    ${retry_interval}
    ...    Check Element Visibility Safely
    ...    ${locator}

Wait For Rater Status Complete
    [Arguments]    ${locator}    ${timeout}=600s    ${interval}=5s
    Safe Wait Until Element Is Visible    ${locator}    120s
    WHILE    True    limit=${timeout}
        ${status_text}=    Set Variable    ${EMPTY}
        ${res_call}    ${res}=    Run Keyword And Ignore Error    RPA.Browser.Selenium.Get Text    ${locator}
        IF    '${res_call}' == 'PASS' and $res is not None and '${res}' != ''
            ${status_text}=    Set Variable    ${res}
        END
        IF    '${status_text}' == ''
            ${res_call}    ${res}=    Run Keyword And Ignore Error    RPA.Browser.Selenium.Get Element Attribute    ${locator}    textContent
            IF    '${res_call}' == 'PASS' and $res is not None and '${res}' != ''
                ${status_text}=    Set Variable    ${res}
            END
        END
        IF    '${status_text}' == ''
            ${res_call}    ${res}=    Run Keyword And Ignore Error    RPA.Browser.Selenium.Get Value    ${locator}
            IF    '${res_call}' == 'PASS' and $res is not None and '${res}' != ''
                ${status_text}=    Set Variable    ${res}
            END
        END
        ${status_str}=    Convert To String    ${status_text}
        ${status_clean}=    Strip String    ${status_str}
        ${status_lower}=    Convert To Lower Case    ${status_clean}
        IF    'complete' in $status_lower
            Log    Rater status completed successfully: '${status_clean}'
            BREAK
        ELSE IF    'in progress' in $status_lower or 'in pregress' in $status_lower or 'inprogress' in $status_lower or $status_lower == '' or $status_lower == 'none'
            Log    Current rater status: '${status_clean}'. Waiting ${interval}...
            Sleep    ${interval}
        ELSE
            Fail    Automation stopped: Unexpected rater status '${status_clean}' encountered (expected 'Complete' or 'In Progress').
        END
    END

Extract Risk Number
    [Arguments]    ${flow_name}
    Safe Wait Until Element Is Visible    //*[@id="ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlRiskHeader_btnReturnToRiskSummary"]    15s
    ${risk_number}=    RPA.Browser.Selenium.Get Text    //*[@id="ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlRiskHeader_btnReturnToRiskSummary"]
    Evaluate    open(r'${CURDIR}${/}risk_numbers.txt', 'a').write('${flow_name} - Risk Number: ${risk_number}\\n')

Enter Terror All Other 100
    [Arguments]    ${locator}=(//input[contains(@id, 'ctrlTria_txtAllOther') or contains(@id, 'txtAllOther')])[1]
    Safe Wait Until Element Is Visible    ${locator}    120s
    Execute Javascript    window.scrollTo(0, document.body.scrollHeight)
    Sleep    1s
    Safe Click Element    ${locator}
    Press Keys    ${locator}    CTRL+a+BACKSPACE
    Run Keyword And Ignore Error    Clear Element Text    ${locator}
    Press Keys    ${locator}    100
    Execute Javascript
    ...    var el = document.evaluate("(//input[contains(@id, 'ctrlTria_txtAllOther') or contains(@id, 'txtAllOther')])[1]", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
    ...    if (el) {
    ...        el.value = '100';
    ...        el.dispatchEvent(new Event('input', { bubbles: true }));
    ...        el.dispatchEvent(new Event('change', { bubbles: true }));
    ...    }
    Press Keys    ${locator}    TAB
    Sleep    2s
    ${val}=    RPA.Browser.Selenium.Get Value    ${locator}
    ${val_str}=    Convert To String    ${val}
    IF    '100' not in '${val_str}'
        Execute Javascript
        ...    var el = document.evaluate("(//input[contains(@id, 'ctrlTria_txtAllOther') or contains(@id, 'txtAllOther')])[1]", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
        ...    if (el) {
        ...        el.value = '100';
        ...        el.dispatchEvent(new Event('input', { bubbles: true }));
        ...        el.dispatchEvent(new Event('change', { bubbles: true }));
        ...        el.dispatchEvent(new Event('blur', { bubbles: true }));
        ...    }
        Press Keys    ${locator}    100
        Press Keys    ${locator}    TAB
    END
