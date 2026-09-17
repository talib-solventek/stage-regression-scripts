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
E&S US Prop Toronto
    Open Workbook    ${DATA_SOURCE}
    ${excel_rows}=    Read Worksheet    E&S US Prop Toronto
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
    ...    Submission (Pricing Completed)
	Copy Risk
    Sandbox Quote
    Manage Sandbox
    Finalize Tech Prem
    View Risk
    Close Browser

Login to App
    ${URL}=    Get From Dictionary    ${current_row}    A
    ${NUSER}=    Get From Dictionary    ${current_row}    B
    ${NUSERPASSWORD}=    Get From Dictionary    ${current_row}    C
    Open Available Browser    ${URL}    maximized=${TRUE}    browser_selection=edge
    Press Keys    None    CTRL+R
    Set Selenium Implicit Wait    30s
    Click Element When Visible    xpath=//a[normalize-space()='Create New Risk >']
    Click Element When Visible    xpath=(//a[contains(text(),'US E&S Property')])[2]
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
    ...    (//span[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_riskHeader_labelStatus'])[1]
    ...    Submission
    Fill Pricing Details[First Run]

Fill Insured Details[First Run]
    Click Element When Visible    (//span[normalize-space()='No'])[1]
    Sleep    2s
    Click Element When Visible    (//span[normalize-space()='No'])[3]
    Click Element When Visible    (//span[normalize-space()='No'])[4]
    ${iname}=    Get From Dictionary    ${current_row}    D
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_TextBoxFirstNamedInsured1'])[1]
    ...    ${iname}
    ${state}=    Get From Dictionary    ${current_row}    E
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_ddUSState'])[1]
    ...    ${state}
    ${city}=    Get From Dictionary    ${current_row}    F
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_ddUSCity'])[1]
    ...    ${city}
    ${adl1}=    Get From Dictionary    ${current_row}    G
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_textBoxUSAddressLine1'])[1]
    ...    ${adl1}
    ${zip1}=    Get From Dictionary    ${current_row}    H
        Click Element When Visible
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_textBoxUSZipCode'])[1]
        ${zipelm}=    Get WebElement
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_textBoxUSZipCode'])[1]
        ${attribute}=    Get Element Attribute
        ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_textBoxUSZipCode'])[1]
        ...    value
        Input Text
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
    Sleep    10s
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
    ...    Toronto
	Sleep	5s
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListCompany'])[1]
    ...    CA-CBC
    Select From List By Index
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListUnderWriter'])[1]
    ...    1
    Select From List By Index
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListAssistant'])[1]
    ...    1
    ${industry_code}=    Get From Dictionary    ${current_row}    K
    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_TextBoxIndustryCode'])[1]
    ...    ${industry_code}
    ${industry_code_full}=    Get From Dictionary    ${current_row}    K
    Click Element When Visible    (//a[normalize-space()='${industry_code_full}'])[1]
	 Select From List By Index
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListDepartment'])[1]
    ...    2
    ${effdate}=    Get From Dictionary    ${current_row}    J
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DatePickerEffectiveDate_textDate'])[1]
    ...    ${effdate}
   	${submission_tiv}=    Get From Dictionary    ${current_row}    J
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_txtSubmissionTiv'])[1]
    ...    10000
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//legend[normalize-space()='Broker Info'])[1]    20s
    ${brokfirm}=    Get From Dictionary    ${current_row}    M
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
    ${brokcont}=    Get From Dictionary    ${current_row}    N
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
    Wait Until Element Is Visible    (//a[normalize-space()='Contract I + Location'])[1]    20s
    Choose File
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlLayerSubLimit_fileuploadAirContratLayer'])[1]
    ...    ${CURDIR}${/}Contract.csv
    Sleep    60s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlLayerSubLimit_UploadAirContractLayer'])[1]
    Sleep    5s
    Choose File
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlLayerSubLimit_locationUpload'])[1]
    ...    ${CURDIR}${/}Location.csv
    Sleep    60s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlLayerSubLimit_btnUpload'])[1]
    Sleep    5s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='Contract II'])[1]    10s
    ${limit}=    Get From Dictionary    ${current_row}    O
    Input Text When Element Is Visible
    ...    (//input[contains(@id,'_rptAirContractIAccountInformation_ctl05_TextBoxLimit')])
    ...    25M
    ${attachment}=    Get From Dictionary    ${current_row}    P
    Input Text When Element Is Visible
    ...    (//input[contains(@id,'_rptAirContractIAccountInformation_ctl05_TextBoxAttachment')])
    ...    ${attachment}
    ${deduction}=    Get From Dictionary    ${current_row}    Q
    Input Text When Element Is Visible    (//input[contains(@id,'TextBoxAOPDed')])    ${deduction}
    Click Element When Visible    (//input[contains(@id,'RptAirContractSSInput_ctl01_addLayer')])
    ${ss_limit}=    Get From Dictionary    ${current_row}    R
    Input Text    (//input[contains(@id,'RptAirContractSSInput_ctl01_TextBoxSSLimit')])    ${ss_limit}
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='AIR Results Occurence'])[1]    20s
    FOR    ${counter}    IN RANGE    1    7
        ${next_button_visible}=    Is Element Visible
        ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_btnNext'])[1]
        IF    ${next_button_visible} == True
            Click Element When Visible
            ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_btnNext'])[1]
            BREAK
        END
        Sleep    100s
    END
    Wait Until Element Is Visible    (//a[normalize-space()='Pricing'])[1]    20s
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='Yes'])[1]
    ${layer_limit}=    Get From Dictionary    ${current_row}    T
    Input Text    (//input[contains(@id,'textBoxLayerLimit')])    ${layer_limit}
    ${Xs_of}=    Get From Dictionary    ${current_row}    U
    Input Text    (//input[contains(@id,'textBoxXSof')])    ${Xs_of}
    ${ironshore_limit}=    Get From Dictionary    ${current_row}    V
    Input Text    (//input[contains(@id,'textBoxIronshoreGross')])    ${ironshore_limit}
    ${layer_premium}=    Get From Dictionary    ${current_row}    W
    Input Text    (//input[contains(@id,'textBoxLayerPremium')])    ${layer_premium}
    ${layer_premium_tria}=    Get From Dictionary    ${current_row}    X
    Input Text    (//input[contains(@id,'textBoxLayerPremiumTRIA')])    ${layer_premium_tria}
    ${commission}=    Get From Dictionary    ${current_row}    Y
    Input Text    (//input[contains(@id,'_textBoxBrokerCommisionPercentage')])    ${commission}
    ${layer_losses_1}=    Get From Dictionary    ${current_row}    Z
    Input Text    (//input[contains(@id,'LayerLosses')])[1]    ${layer_losses_1}
    ${layer_losses_2}=    Get From Dictionary    ${current_row}    AA
    Input Text    (//input[contains(@id,'LayerLosses')])[2]    ${layer_losses_2}
    ${layer_losses_3}=    Get From Dictionary    ${current_row}    AB
    Input Text    (//input[contains(@id,'LayerLosses')])[3]    ${layer_losses_3}
    ${layer_losses_4}=    Get From Dictionary    ${current_row}    AC
    Input Text    (//input[contains(@id,'LayerLosses')])[4]    ${layer_losses_4}
    ${layer_losses_5}=    Get From Dictionary    ${current_row}    AD
    Input Text    (//input[contains(@id,'LayerLosses')])[5]    ${layer_losses_5}
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='Yes'])[4]
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='Yes'])[5]
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[2]
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[3]
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]

Fill Insured Details[Copy Risk]
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[1]
    Sleep    2s
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[3]
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[4]
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
    ...    Toronto
    Select From List By Label
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListCompany'])[1]
    ...    CA-CBC
    Select From List By Index
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListUnderWriter'])[1]
    ...    1
    Select From List By Index
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListAssistant'])[1]
    ...    1
    ${industry_code}=    Get From Dictionary    ${current_row}    K
    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_TextBoxIndustryCode'])[1]
    ...    ${industry_code}
    ${industry_code_full}=    Get From Dictionary    ${current_row}    K
    Click Element When Visible    (//a[normalize-space()='${industry_code_full}'])[1]
	 Select From List By Index
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListDepartment'])[1]
    ...    2
    ${effdate}=    Get From Dictionary    ${current_row}    J
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DatePickerEffectiveDate_textDate'])[1]
    ...    ${effdate}
   	${submission_tiv}=    Get From Dictionary    ${current_row}    J
    Input Text
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_txtSubmissionTiv'])[1]
    ...    10000
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]
    Wait Until Element Is Visible    (//legend[normalize-space()='Broker Info'])[1]    20s
    ${brokfirm}=    Get From Dictionary    ${current_row}    M
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
    ${brokcont}=    Get From Dictionary    ${current_row}    N
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

Fill Pricing Details[Copy Risk]
    Wait Until Element Is Visible    (//a[normalize-space()='Contract I + Location'])[1]    20s
    Choose File
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlLayerSubLimit_fileuploadAirContratLayer'])[1]
    ...    ${CURDIR}${/}Contract.csv
    Sleep    60s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlLayerSubLimit_UploadAirContractLayer'])[1]
    Sleep    5s
    Choose File
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlLayerSubLimit_locationUpload'])[1]
    ...    ${CURDIR}${/}Location.csv
    Sleep    60s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlLayerSubLimit_btnUpload'])[1]
    Sleep    5s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]
   Wait Until Element Is Visible    (//a[normalize-space()='Contract II'])[1]    10s
    ${limit}=    Get From Dictionary    ${current_row}    O
    Input Text When Element Is Visible
    ...    (//input[contains(@id,'_rptAirContractIAccountInformation_ctl05_TextBoxLimit')])
    ...    25M
    ${attachment}=    Get From Dictionary    ${current_row}    P
    Input Text When Element Is Visible
    ...    (//input[contains(@id,'_rptAirContractIAccountInformation_ctl05_TextBoxAttachment')])
    ...    ${attachment}
    ${deduction}=    Get From Dictionary    ${current_row}    Q
    Input Text When Element Is Visible    (//input[contains(@id,'TextBoxAOPDed')])    ${deduction}
    Click Element When Visible    (//input[contains(@id,'RptAirContractSSInput_ctl01_addLayer')])
    ${ss_limit}=    Get From Dictionary    ${current_row}    R
    Input Text    (//input[contains(@id,'RptAirContractSSInput_ctl01_TextBoxSSLimit')])    ${ss_limit}
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='AIR Results Occurence'])[1]    20s
    FOR    ${counter}    IN RANGE    1    7
        ${next_button_visible}=    Is Element Visible
        ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_btnNext'])[1]
        IF    ${next_button_visible} == True
            Click Element When Visible
            ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_btnNext'])[1]
            BREAK
        END
        Sleep    100s
    END
     Wait Until Element Is Visible    (//a[normalize-space()='Pricing'])[1]    20s
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='Yes'])[1]
    ${layer_limit}=    Get From Dictionary    ${current_row}    T
    Input Text    (//input[contains(@id,'textBoxLayerLimit')])    ${layer_limit}
    ${Xs_of}=    Get From Dictionary    ${current_row}    U
    Input Text    (//input[contains(@id,'textBoxXSof')])    ${Xs_of}
    ${ironshore_limit}=    Get From Dictionary    ${current_row}    V
    Input Text    (//input[contains(@id,'textBoxIronshoreGross')])    ${ironshore_limit}
    ${layer_premium}=    Get From Dictionary    ${current_row}    W
    Input Text    (//input[contains(@id,'textBoxLayerPremium')])    ${layer_premium}
    ${layer_premium_tria}=    Get From Dictionary    ${current_row}    X
    Input Text    (//input[contains(@id,'textBoxLayerPremiumTRIA')])    ${layer_premium_tria}
    ${commission}=    Get From Dictionary    ${current_row}    Y
    Input Text    (//input[contains(@id,'_textBoxBrokerCommisionPercentage')])    ${commission}
    ${layer_losses_1}=    Get From Dictionary    ${current_row}    Z
    Input Text    (//input[contains(@id,'LayerLosses')])[1]    ${layer_losses_1}
    ${layer_losses_2}=    Get From Dictionary    ${current_row}    AA
    Input Text    (//input[contains(@id,'LayerLosses')])[2]    ${layer_losses_2}
    ${layer_losses_3}=    Get From Dictionary    ${current_row}    AB
    Input Text    (//input[contains(@id,'LayerLosses')])[3]    ${layer_losses_3}
    ${layer_losses_4}=    Get From Dictionary    ${current_row}    AC
    Input Text    (//input[contains(@id,'LayerLosses')])[4]    ${layer_losses_4}
    ${layer_losses_5}=    Get From Dictionary    ${current_row}    AD
    Input Text    (//input[contains(@id,'LayerLosses')])[5]    ${layer_losses_5}
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='Yes'])[4]
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='Yes'])[5]
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[2]
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[3]
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]

Fill Pricing Details[Sandbox]
	Wait Until Element Is Visible    (//a[normalize-space()='Contract I + Location'])[1]    20s
    Choose File
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlLayerSubLimit_fileuploadAirContratLayer'])[1]
    ...    ${CURDIR}${/}Contract.csv
    Sleep    60s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlLayerSubLimit_UploadAirContractLayer'])[1]
    Sleep    5s
    Choose File
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlLayerSubLimit_locationUpload'])[1]
    ...    ${CURDIR}${/}Location.csv
    Sleep    60s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlLayerSubLimit_btnUpload'])[1]
    Sleep    5s
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]
   Wait Until Element Is Visible    (//a[normalize-space()='Contract II'])[1]    10s
    ${limit}=    Get From Dictionary    ${current_row}    O
    Input Text When Element Is Visible
    ...    (//input[contains(@id,'_rptAirContractIAccountInformation_ctl05_TextBoxLimit')])
    ...    25M
    ${attachment}=    Get From Dictionary    ${current_row}    P
    Input Text When Element Is Visible
    ...    (//input[contains(@id,'_rptAirContractIAccountInformation_ctl05_TextBoxAttachment')])
    ...    ${attachment}
    ${deduction}=    Get From Dictionary    ${current_row}    Q
    Input Text When Element Is Visible    (//input[contains(@id,'TextBoxAOPDed')])    ${deduction}
    Click Element When Visible    (//input[contains(@id,'RptAirContractSSInput_ctl01_addLayer')])
    ${ss_limit}=    Get From Dictionary    ${current_row}    R
    Input Text    (//input[contains(@id,'RptAirContractSSInput_ctl01_TextBoxSSLimit')])    ${ss_limit}
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]
    Wait Until Element Is Visible    (//a[normalize-space()='AIR Results Occurence'])[1]    20s
    FOR    ${counter}    IN RANGE    1    7
        ${next_button_visible}=    Is Element Visible
        ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_btnNext'])[1]
        IF    ${next_button_visible} == True
            Click Element When Visible
            ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_btnNext'])[1]
            BREAK
        END
        Sleep    100s
    END
     Wait Until Element Is Visible    (//a[normalize-space()='Pricing'])[1]    20s
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='Yes'])[1]
    ${layer_limit}=    Get From Dictionary    ${current_row}    T
    Input Text    (//input[contains(@id,'textBoxLayerLimit')])    ${layer_limit}
    ${Xs_of}=    Get From Dictionary    ${current_row}    U
    Input Text    (//input[contains(@id,'textBoxXSof')])    ${Xs_of}
    ${ironshore_limit}=    Get From Dictionary    ${current_row}    V
    Input Text    (//input[contains(@id,'textBoxIronshoreGross')])    ${ironshore_limit}
    ${layer_premium}=    Get From Dictionary    ${current_row}    W
    Input Text    (//input[contains(@id,'textBoxLayerPremium')])    ${layer_premium}
    ${layer_premium_tria}=    Get From Dictionary    ${current_row}    X
    Input Text    (//input[contains(@id,'textBoxLayerPremiumTRIA')])    ${layer_premium_tria}
    ${commission}=    Get From Dictionary    ${current_row}    Y
    Input Text    (//input[contains(@id,'_textBoxBrokerCommisionPercentage')])    ${commission}
    ${layer_losses_1}=    Get From Dictionary    ${current_row}    Z
    Input Text    (//input[contains(@id,'LayerLosses')])[1]    ${layer_losses_1}
    ${layer_losses_2}=    Get From Dictionary    ${current_row}    AA
    Input Text    (//input[contains(@id,'LayerLosses')])[2]    ${layer_losses_2}
    ${layer_losses_3}=    Get From Dictionary    ${current_row}    AB
    Input Text    (//input[contains(@id,'LayerLosses')])[3]    ${layer_losses_3}
    ${layer_losses_4}=    Get From Dictionary    ${current_row}    AC
    Input Text    (//input[contains(@id,'LayerLosses')])[4]    ${layer_losses_4}
    ${layer_losses_5}=    Get From Dictionary    ${current_row}    AD
    Input Text    (//input[contains(@id,'LayerLosses')])[5]    ${layer_losses_5}
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='Yes'])[4]
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='Yes'])[5]
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[2]
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[3]
    Click Element When Visible
    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]

View Risk
    Wait Until Element Is Visible    (//a[normalize-space()='View Risk'])[1]    30s
    Click Element When Visible    (//a[normalize-space()='View Risk'])[1]
    Sleep    5s
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonViewCurrent'])[1]
    Sleep    5s
    Verify Status    (//legend[@id='fsMainContentLegend'])[1]    View Risk

Copy Risk
    Wait Until Element Is Visible    (//a[normalize-space()='Copy Risk'])[1]    30s
    Click Element When Visible    (//a[normalize-space()='Copy Risk'])[1]
    Sleep    10s
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='Yes'])[3]
	Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='Yes'])[4]
    Fill Insured Details[Copy Risk]
    Fill Pricing Details[Copy Risk]

Sandbox Quote
    Wait Until Element Is Visible    (//a[normalize-space()='Sandbox Quote'])[1]    30s
    Click Element When Visible    (//a[normalize-space()='Sandbox Quote'])[1]
    Sleep    5s
    ${rate_monitor_coment}=    Get From Dictionary    ${current_row}    AR
    Input Text When Element Is Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_repeaterQuoteOptions_ctl01_quoteTab_txtSandboxTitle'])[1]
    ...    test
    Click Element When Visible    (//span[@class='ui-button-text'][normalize-space()='Yes'])[12]
    Fill Pricing Details[Sandbox]

Manage Sandbox
	Wait Until Element Is Visible    (//legend[@id='fsMainContentLegend'])[1]    30s
    ${Reference}=    Get From Dictionary    ${current_row}    AR
    Input Text When Element Is Visible	(//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_txtReferenceRisk'])[1]	8110168992-01
	Sleep	10s
	Select From List By Index
    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ddlQuoteOptionReference'])[1]
    ...    1
	${Sandbox_title}=    Get From Dictionary    ${current_row}    AR
    Input Text When Element Is Visible	(//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_txtSandboxTitle'])[1]	test
    Click Element When Visible    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_BtnSave'])[1]
    Click Element When Visible
    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ButtonReturn'])[1]

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
