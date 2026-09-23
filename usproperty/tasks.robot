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

${EDGE_DRIVER}           ${CURDIR}${/}msedgedriver.exe


*** Tasks ***

US PROPERTY

    Remove File    ${CURDIR}${/}risk_number.txt

    Open Workbook    ${DATA_SOURCE}

    ${excel_rows}=    Read Worksheet    USProperty

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
        
        BREAK

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


    Press Keys    None    CTRL+R

    Set Selenium Implicit Wait    30s

    

    Set Selenium Implicit Wait    30s

    Safe Click Element    xpath=//a[normalize-space()='Create New Risk >']

    Safe Click Element    xpath=(//a[contains(text(),'US E&S Property')])[2]

    Safe Click Element

    ...    id:ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlClearanceSearch_CreateInsured

    ${EMAIL}=    Set Variable    ${NUSER}@libertymutual.com

    Wait Until Element Is Visible    xpath=//input[@type='email']    30s

    Clear Element Text    xpath=//input[@type='email']

    Safe Input Text    xpath=//input[@type='email']    ${EMAIL}

    Press Keys    xpath=//input[@type='email']    ENTER

    Wait Until Element Is Visible    xpath=//input[@type='password']    30s

    Clear Element Text    xpath=//input[@type='password']

    Safe Input Text    xpath=//input[@type='password']    ${NUSERPASSWORD}

    Press Keys    xpath=//input[@type='password']    ENTER


Risk Creation[First Run]

    Fill Insured Details[First Run]

    Verify Status

    ...    xpath=//span[contains(@id, 'labelStatus')]

    ...    Submission

    Fill Pricing Details[First Run]


Fill Insured Details[First Run]

    Store Risk Number

    Safe Click Element    (//span[normalize-space()='No'])[1]

    Safe Click Element    (//span[normalize-space()='No'])[3]

    Safe Click Element    (//span[normalize-space()='No'])[4]

    ${iname}=    Get From Dictionary    ${current_row}    D

    Safe Input Text

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_TextBoxFirstNamedInsured1'])[1]

    ...    ${iname}

    ${state}=    Get From Dictionary    ${current_row}    E

    Safe Select From List By Label

    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_ddUSState'])[1]

    ...    ${state}

    ${city}=    Get From Dictionary    ${current_row}    F

    Safe Select From List By Label

    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_ddUSCity'])[1]

    ...    ${city}

    ${adl1}=    Get From Dictionary    ${current_row}    G

    Safe Input Text

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_textBoxUSAddressLine1'])[1]

    ...    ${adl1}

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

    ${pobox}=    Get From Dictionary    ${current_row}    I

    Safe Input Text

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_FirstNamedInsured_AddressInsured_textBoxUSPoBox'])[1]

    ...    ${pobox}

    Safe Click Element

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]

    Wait Until Element Is Visible

    ...    (//legend[@class='ui-widget ui-widget-header ui-corner-all'][normalize-space()='Additional Named Insured'])[1]

    ...    20s

    Safe Click Element

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]

    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]    20s

    Wait Until Element Is Visible    //*[@id="ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlRiskHeader_btnReturnToRiskSummary"]    15s
    ${risk_number}=    RPA.Browser.Selenium.Get Text    //*[@id="ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlRiskHeader_btnReturnToRiskSummary"]
    Evaluate    open(r'${CURDIR}${/}risk_numbers.txt', 'a').write('First Run - Risk Number: ${risk_number}\\n')

    Safe Select From List By Label

    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListBranch'])[1]

    ...    Atlanta Retail

    Safe Select From List By Label

    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListCompany'])[1]

    ...    LSI2

    Safe Select From List By Index

    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListUnderWriter'])[1]

    ...    1

    Safe Select From List By Index

    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListAssistant'])[1]

    ...    1

    ${industry_code}=    Get From Dictionary    ${current_row}    K

    Safe Input Text

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_TextBoxIndustryCode'])[1]

    ...    ${industry_code}

    ${industry_code_full}=    Get From Dictionary    ${current_row}    K

    Safe Click Element    (//a[normalize-space()='${industry_code_full}'])[1]

    ${effdate}=    Get From Dictionary    ${current_row}    J

    Safe Input Text

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DatePickerEffectiveDate_textDate'])[1]

    ...    ${effdate}

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

    Wait Until Element Is Visible    xpath=//table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']/tbody/tr/td//input[@type="submit"] | //table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']/tr/td//input[@type="submit"]    30s

    ${table_elements}=    Get WebElements

    ...    xpath=//table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']/tbody/tr/td//input[@type="submit"] | //table[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerFirmTableBrokerFirm']/tr/td//input[@type="submit"]

    Log    Found ${table_elements.__len__()} elements in the broker info table

    Safe Click Element    ${table_elements}[2]

    Press Keys    None    PAGE_DOWN

    Wait Until Keyword Succeeds

    ...    3x

    ...    5s

    ...    Safe Click Element

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]

    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]    20s

    ${brokcont}=    Get From Dictionary    ${current_row}    N

    Safe Input Text

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_ctrlBrokerContactSearch_TextBoxBrokerContact'])[1]

    ...    ${brokcont}

    Safe Click Element

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_ctrlBrokerContactSearch_ButtonBrokerContactSearchStartsWith'])[1]

    Wait Until Element Is Visible    xpath=//table[@id='TableBrokerContactSearch']/tbody/tr/td//input[@value="Select"] | //table[@id='TableBrokerContactSearch']/tr/td//input[@value="Select"]    30s

    ${table_elements}=    Get WebElements

    ...    xpath=//table[@id='TableBrokerContactSearch']/tbody/tr/td//input[@value="Select"] | //table[@id='TableBrokerContactSearch']/tr/td//input[@value="Select"]

    Log    Found ${table_elements.__len__()} elements in the broker contact table

    Safe Click Element    ${table_elements}[0]

    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]

    Safe Click Element

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_ButtonSave'])[1]

    Press Keys    None    PAGE_DOWN

    ${yes_visible}=    Is Element Visible    xpath=//*[@id="ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContactSurplus_rblSurplusLinesBroker"]/label[1]/span

    IF    ${yes_visible} == True

        Safe Click Element    xpath=//*[@id="ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContactSurplus_rblSurplusLinesBroker"]/label[1]/span

    END

    Wait Until Keyword Succeeds

    ...    3x

    ...    5s

    ...    Safe Click Element

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]

    ${continue_visible}=    Is Element Visible    xpath=/html/body/div[3]/div[11]/div/button[2]/span

    IF    ${continue_visible} == True

        Safe Click Element    xpath=/html/body/div[3]/div[11]/div/button[2]/span

    END


Fill Pricing Details[First Run]

    Wait Until Element Is Visible    (//a[normalize-space()='Contract I + Location'])[1]    20s

    Safe Choose File

    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlLayerSubLimit_fileuploadAirContratLayer'])[1]

    ...    ${CURDIR}${/}Contract.csv

    Sleep    60s

    Safe Click Element

    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlLayerSubLimit_UploadAirContractLayer'])[1]

    Safe Choose File

    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlLayerSubLimit_locationUpload'])[1]

    ...    ${CURDIR}${/}Location.csv

    Sleep    60s

    Safe Click Element

    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlLayerSubLimit_btnUpload'])[1]

    Safe Click Element

    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]

    Wait Until Element Is Visible    (//a[normalize-space()='Contract II'])[1]    10s

    ${limit}=    Get From Dictionary    ${current_row}    O

    Safe Input Text

    ...    (//input[contains(@id,'_rptAirContractIAccountInformation_ctl05_TextBoxLimit')])

    ...    25M

    ${attachment}=    Get From Dictionary    ${current_row}    P

    Safe Input Text

    ...    (//input[contains(@id,'_rptAirContractIAccountInformation_ctl05_TextBoxAttachment')])

    ...    ${attachment}

    ${deduction}=    Get From Dictionary    ${current_row}    Q

    Safe Input Text    (//input[contains(@id,'TextBoxAOPDed')])    ${deduction}

    Safe Click Element    (//input[contains(@id,'RptAirContractSSInput_ctl01_addLayer')])

    ${ss_limit}=    Get From Dictionary    ${current_row}    R

    Safe Input Text    (//input[contains(@id,'RptAirContractSSInput_ctl01_TextBoxSSLimit')])    ${ss_limit}

    Safe Click Element

    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]

    Wait Until Element Is Visible    (//a[normalize-space()='AIR Results Occurence'])[1]    20s

    FOR    ${counter}    IN RANGE    1    7

        ${next_button_visible}=    Is Element Visible

        ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_btnNext'])[1]

        IF    ${next_button_visible} == True

            Safe Click Element

            ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_btnNext'])[1]

            BREAK

        END

        Sleep    250s

    END

    Wait Until Element Is Visible    (//a[normalize-space()='Pricing'])[1]    20s

    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Yes'])[1]

    ${layer_limit}=    Get From Dictionary    ${current_row}    T

    Safe Input Text    (//input[contains(@id,'textBoxLayerLimit')])    ${layer_limit}

    ${Xs_of}=    Get From Dictionary    ${current_row}    U

    Safe Input Text    (//input[contains(@id,'textBoxXSof')])    ${Xs_of}

    ${ironshore_limit}=    Get From Dictionary    ${current_row}    V

    Safe Input Text    (//input[contains(@id,'textBoxIronshoreGross')])    ${ironshore_limit}

    ${layer_premium}=    Get From Dictionary    ${current_row}    W

    Safe Input Text    (//input[contains(@id,'textBoxLayerPremium')])    ${layer_premium}

    ${layer_premium_tria}=    Get From Dictionary    ${current_row}    X

    Safe Input Text    (//input[contains(@id,'textBoxLayerPremiumTRIA')])    ${layer_premium_tria}

    ${commission}=    Get From Dictionary    ${current_row}    Y

    Safe Input Text    (//input[contains(@id,'_textBoxBrokerCommisionPercentage')])    ${commission}

    ${layer_losses_1}=    Get From Dictionary    ${current_row}    Z

    Safe Input Text    (//input[contains(@id,'LayerLosses')])[1]    ${layer_losses_1}

    ${layer_losses_2}=    Get From Dictionary    ${current_row}    AA

    Safe Input Text    (//input[contains(@id,'LayerLosses')])[2]    ${layer_losses_2}

    ${layer_losses_3}=    Get From Dictionary    ${current_row}    AB

    Safe Input Text    (//input[contains(@id,'LayerLosses')])[3]    ${layer_losses_3}

    ${layer_losses_4}=    Get From Dictionary    ${current_row}    AC

    Safe Input Text    (//input[contains(@id,'LayerLosses')])[4]    ${layer_losses_4}

    ${layer_losses_5}=    Get From Dictionary    ${current_row}    AD

    Safe Input Text    (//input[contains(@id,'LayerLosses')])[5]    ${layer_losses_5}

    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Yes'])[4]

    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Yes'])[5]

    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[2]

    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[3]

    Safe Click Element

    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]

    Wait Until Element Is Visible    (//a[normalize-space()='Renewal Comparison'])[1]    20s

    Safe Click Element

    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderCenter$ctl00'])[1]

    Wait Until Element Is Visible    (//a[normalize-space()='Deductible And Coverage'])[1]    20s

    ${policy_form}=    Get From Dictionary    ${current_row}    AE

    Safe Select From List By Label

    ...    (//select[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlDeductibleAndCoverages_ddlPolicyForm'])[1]

    ...    ${policy_form}

    ${peril_insured}=    Get From Dictionary    ${current_row}    AF

    Safe Select From List By Label

    ...    (//select[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlDeductibleAndCoverages_ddlPerilInsured'])[1]

    ...    ${peril_insured}

    ${property_covered}=    Get From Dictionary    ${current_row}    AG

    Safe Select From List By Label

    ...    (//select[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlDeductibleAndCoverages_ddlPropertyCovered'])[1]

    ...    ${property_covered}

    ${comments}=    Get From Dictionary    ${current_row}    AH

    Safe Input Text

    ...    (//textarea[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlDeductibleAndCoverages_textBoxPleaseSpecifyPC'])[1]

    ...    ${comments}

    Safe Click Element

    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]

    Wait Until Element Is Visible    (//a[normalize-space()='Quote Items'])[1]    20s

    ${min_earned_prem}=    Get From Dictionary    ${current_row}    AI

    Safe Input Text

    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlQuoteItems_textBoxMiniumEarnedPremium'])[1]

    ...    ${min_earned_prem}

    ${valuation_property}=    Get From Dictionary    ${current_row}    AJ

    Safe Select From List By Label

    ...    (//select[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlQuoteItems_ddlValuation'])[1]

    ...    ${valuation_property}

    ${business_interuption}=    Get From Dictionary    ${current_row}    AK

    Safe Select From List By Label

    ...    (//select[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlQuoteItems_ddlValuationBusinessInterruption'])[1]

    ...    ${business_interuption}

    ${coinsurance_property}=    Get From Dictionary    ${current_row}    AL

    Safe Input Text

    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlQuoteItems_txtCoinsuranceProperty'])[1]

    ...    ${coinsurance_property}

    ${coinsurance_business}=    Get From Dictionary    ${current_row}    AM

    Safe Input Text

    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlQuoteItems_txtCoinsuranceBusiness'])[1]

    ...    ${coinsurance_business}

    ${perils}=    Get From Dictionary    ${current_row}    AN

    Safe Input Text

    ...    (//textarea[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlQuoteItems_txtPerils'])[1]

    ...    ${perils}

    Safe Click Element

    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]

    Wait Until Element Is Visible    (//a[normalize-space()='Additional Terms And Conditions'])[1]    20s

    ${ocurrence_limit}=    Get From Dictionary    ${current_row}    AO

    Safe Select From List By Label

    ...    (//select[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlTermsAndConditions_ddlOccurrenceLimit'])[1]

    ...    ${ocurrence_limit}

    Safe Click Element

    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]


Pre Bind Endorsement

    Safe Click Element    (//input[@id='btnShowAllEndorsementTypes'])[1]

    ${endorsement_type}=    Get From Dictionary    ${current_row}    AP

    Safe Input Text

    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlPreBindEndorsements_txtEndorsementType'])[1]

    ...    ${endorsement_type}

    Safe Click Element    (//a[normalize-space()='${endorsement_type}'])[1]

    Safe Click Element

    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlPreBindEndorsements_btAddEndorsement'])[1]

    Safe Click Element

    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]

    Wait Until Keyword Succeeds    3x    5s    Wait Until Element Is Visible    (//a[normalize-space()='Subjectivities'])[1]    2s

    Safe Click Element

    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]


Quote

    Wait Until Element Is Visible    (//a[normalize-space()='Finalize'])[1]    20s

    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Yes'])[1]

    Safe Click Element    (//span[normalize-space()='Correct'])[1]

    ${comments_under_quote}=    Get From Dictionary    ${current_row}    AR

    Safe Input Text

    ...    (//textarea[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlFinalize_txtComments'])[1]

    ...    ${comments_under_quote}

    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[3]

    Safe Click Element

    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_btnGenerateQuote'])[1]


Ready To Bind

    Wait Until Element Is Visible    (//a[normalize-space()='Account Summary'])[1]    30s

    Safe Click Element    (//a[normalize-space()='Account Summary'])[1]

    ${comments_under_bind}=    Get From Dictionary    ${current_row}    AR

    Safe Input Text

    ...    (//textarea[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlAccountSummary_textBoxReferralComments'])[1]

    ...    ${comments_under_bind}

    Safe Select From List By Index

    ...    (//select[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlAccountSummary_ddlApprover'])[1]

    ...    1

    Safe Click Element

    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_btnComplete'])[1]

    Safe Click Element    (//a[normalize-space()='Bind'])[1]

    Wait Until Element Is Visible    (//legend[@id='fsMainContentLegend'])[1]    30s

    Safe Click Element

    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl01'])[1]

    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]    30s

    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Yes'])[1]

    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[2]

    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[3]

    Safe Click Element

    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_btnBind'])[1]


Book

    [Arguments]    ${download_policy}

    Wait Until Element Is Visible    (//a[normalize-space()='Book and Issue'])[1]    30s

    Safe Click Element    (//a[normalize-space()='Book and Issue'])[1]

    Wait Until Element Is Visible    (//legend[@id='fsMainContentLegend'])[1]    30s

    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[1]

    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[2]

    Safe Click Element

    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_btnDraft'])[1]


Copy Quote

    Wait Until Element Is Visible    (//a[normalize-space()='Copy Quote'])[1]    30s

    Safe Click Element    (//a[normalize-space()='Copy Quote'])[1]

    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Yes'])[12]

    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Yes'])[13]

    Wait Until Element Is Visible

    ...    xpath=(//span[contains(@id, 'HeaderQuote')])[2]

    ...    10s

    Verify Status

    ...    xpath=(//span[contains(@id, 'HeaderQuote')])[2]

    ...    Option 2 - [Status : Quoted (In Revision)]

    Safe Click Element

    ...    xpath=(//span[contains(@id, 'HeaderQuote')])[1]


Issue Policy

    Wait Until Element Is Visible    (//a[normalize-space()='Issue Policy'])[1]    30s

    Safe Click Element    (//a[normalize-space()='Issue Policy'])[1]

    Safe Click Element

    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]

    ${endorsement_number}=    Get From Dictionary    ${current_row}    AU

    Safe Input Text

    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlPreBindEndorsements_txtEndorsementNo'])[1]

    ...    ${endorsement_number}

    ${endorsement_title}=    Get From Dictionary    ${current_row}    AV

    Safe Input Text

    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlPreBindEndorsements_txtEndorsementTitle'])[1]

    ...    ${endorsement_title}

    Safe Choose File

    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlPreBindEndorsements_fuExternalFormPdf'])[1]

    ...    ${CURDIR}${/}EndorsementBroker.pdf

    Safe Click Element

    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlPreBindEndorsements_btAddAndUploadEndorsement'])[1]

    Safe Click Element

    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]

    Wait Until Element Is Visible    (//legend[@id='fsMainContentLegend'])[1]

    Safe Click Element

    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_btnIssue'])[1]


Verify Status

    [Arguments]    ${location}    ${status_expected}

    Wait Until Keyword Succeeds    10x    3s    Element Text Should Be    ${location}    ${status_expected}


Store Risk Number

    ${status}=    Run Keyword And Return Status

    ...    Wait Until Element Is Visible

    ...    xpath=(//span[contains(@id, 'titlePolicyNumber')])[1]

    ...    20s

    IF    ${status}

        ${risk_number}=    RPA.Browser.Selenium.Get Text    xpath=(//span[contains(@id, 'titlePolicyNumber')])[1]

        ${risk_number}=    Evaluate    "${risk_number}".strip()

        Log    Retrieved Risk Number: ${risk_number}

        Create File    ${CURDIR}${/}risk_number.txt    ${risk_number}    overwrite=${TRUE}

    ELSE

        Log    Warning: Risk number element not found.

    END


Fill Insured Details[Copy Risk]

    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[1]

    Sleep    2s

    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[3]

    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[4]

    Safe Click Element

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]

    Wait Until Element Is Visible

    ...    (//legend[@class='ui-widget ui-widget-header ui-corner-all'][normalize-space()='Additional Named Insured'])[1]

    ...    20s

    Safe Click Element

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]

    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]    20s

    Wait Until Element Is Visible    //*[@id="ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlRiskHeader_btnReturnToRiskSummary"]    15s
    ${risk_number}=    RPA.Browser.Selenium.Get Text    //*[@id="ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlRiskHeader_btnReturnToRiskSummary"]
    Evaluate    open(r'${CURDIR}${/}risk_numbers.txt', 'a').write('Copy Risk - Risk Number: ${risk_number}\\n')

    Safe Select From List By Label

    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListBranch'])[1]

    ...    Atlanta Retail

    Safe Select From List By Label

    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListCompany'])[1]

    ...    ISIC

    Safe Select From List By Index

    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListUnderWriter'])[1]

    ...    1

    Safe Select From List By Index

    ...    (//select[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DropDownListAssistant'])[1]

    ...    1

    ${industry_code}=    Get From Dictionary    ${current_row}    K

    Safe Input Text

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_TextBoxIndustryCode'])[1]

    ...    ${industry_code}

    ${industry_code_full}=    Get From Dictionary    ${current_row}    K

    Safe Click Element    (//a[normalize-space()='${industry_code_full}'])[1]

    ${effdate}=    Get From Dictionary    ${current_row}    J

    Safe Input Text

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_DatePickerEffectiveDate_textDate'])[1]

    ...    ${effdate}

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

    Click Element    ${table_elements}[2]

    Press Keys    None    PAGE_DOWN

    Wait Until Keyword Succeeds

    ...    3x

    ...    5s

    ...    Safe Click Element

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]

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

    Click Element    ${table_elements}[0]

    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]

    Safe Click Element

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_CtrlBrokerContact_ButtonSave'])[1]

    Press Keys    None    PAGE_DOWN

    Wait Until Keyword Succeeds

    ...    3x

    ...    5s

    ...    Safe Click Element

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]


Fill Pricing Details[Copy Risk]

    Wait Until Element Is Visible    (//a[normalize-space()='Contract I + Location'])[1]    20s

    Safe Click Element

    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]

    Wait Until Element Is Visible    (//a[normalize-space()='Contract II'])[1]    10s

    Safe Click Element

    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]

    Wait Until Element Is Visible    (//a[normalize-space()='AIR Results Occurence'])[1]    20s

    Safe Click Element

    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_btnNext'])[1]

    Wait Until Element Is Visible    (//a[normalize-space()='Pricing'])[1]    20s

    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Yes'])[1]

    Safe Click Element

    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]

    Wait Until Element Is Visible    (//a[normalize-space()='Renewal Comparison'])[1]    20s

    Safe Click Element

    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderCenter$ctl00'])[1]

    Wait Until Element Is Visible    (//a[normalize-space()='Deductible And Coverage'])[1]    20s

    Safe Click Element

    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]

    Wait Until Element Is Visible    (//a[normalize-space()='Quote Items'])[1]    20s

    Safe Click Element

    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]

    Wait Until Element Is Visible    (//a[normalize-space()='Additional Terms And Conditions'])[1]    20s

    Safe Click Element

    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]


Risk Creation[Renewal]

    Fill Insured Details[Renewal]

    Verify Status

    ...    xpath=//span[contains(@id, 'labelStatus')]

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

    Wait Until Element Is Visible    //*[@id="ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlRiskHeader_btnReturnToRiskSummary"]    15s
    ${risk_number}=    RPA.Browser.Selenium.Get Text    //*[@id="ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlRiskHeader_btnReturnToRiskSummary"]
    Evaluate    open(r'${CURDIR}${/}risk_numbers.txt', 'a').write('Renewal - Risk Number: ${risk_number}\\n')

    Safe Click Element

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]

    Safe Click Element

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]

    Safe Click Element

    ...    (//input[@id='ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_btnNext'])[1]


Fill Pricing Details[Renewal]

    Wait Until Element Is Visible    (//a[normalize-space()='Contract I + Location'])[1]    20s

    Safe Choose File

    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlLayerSubLimit_fileuploadAirContratLayer'])[1]

    ...    ${CURDIR}${/}Contract.csv

    Sleep    60s

    Safe Click Element

    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlLayerSubLimit_UploadAirContractLayer'])[1]

    Sleep    5s

    Safe Choose File

    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlLayerSubLimit_locationUpload'])[1]

    ...    ${CURDIR}${/}Location.csv

    Sleep    60s

    Safe Click Element

    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlLayerSubLimit_btnUpload'])[1]

    Sleep    5s

    Safe Click Element

    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]

    Wait Until Element Is Visible    (//a[normalize-space()='Contract II'])[1]    10s

    ${limit}=    Get From Dictionary    ${current_row}    O

    Safe Input Text

    ...    (//input[contains(@id,'_rptAirContractIAccountInformation_ctl05_TextBoxLimit')])

    ...    25M

    ${attachment}=    Get From Dictionary    ${current_row}    P

    Safe Input Text

    ...    (//input[contains(@id,'_rptAirContractIAccountInformation_ctl05_TextBoxAttachment')])

    ...    ${attachment}

    ${deduction}=    Get From Dictionary    ${current_row}    Q

    Safe Input Text    (//input[contains(@id,'TextBoxAOPDed')])    ${deduction}

    Safe Click Element    (//input[contains(@id,'RptAirContractSSInput_ctl01_addLayer')])

    ${ss_limit}=    Get From Dictionary    ${current_row}    R

    Safe Input Text    (//input[contains(@id,'RptAirContractSSInput_ctl01_TextBoxSSLimit')])    ${ss_limit}

    Safe Click Element

    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]

    Wait Until Element Is Visible    (//a[normalize-space()='AIR Results Occurence'])[1]    20s

    FOR    ${counter}    IN RANGE    1    7

        ${next_button_visible}=    Is Element Visible

        ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_btnNext'])[1]

        IF    ${next_button_visible} == True

            Safe Click Element

            ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_btnNext'])[1]

            BREAK

        END

        Sleep    100s

    END

    Wait Until Element Is Visible    (//a[normalize-space()='Pricing'])[1]    20s

    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Yes'])[1]

    ${layer_limit}=    Get From Dictionary    ${current_row}    T

    Safe Input Text    (//input[contains(@id,'textBoxLayerLimit')])    ${layer_limit}

    ${Xs_of}=    Get From Dictionary    ${current_row}    U

    Safe Input Text    (//input[contains(@id,'textBoxXSof')])    ${Xs_of}

    ${ironshore_limit}=    Get From Dictionary    ${current_row}    V

    Safe Input Text    (//input[contains(@id,'textBoxIronshoreGross')])    ${ironshore_limit}

    ${layer_premium}=    Get From Dictionary    ${current_row}    W

    Safe Input Text    (//input[contains(@id,'textBoxLayerPremium')])    ${layer_premium}

    ${layer_premium_tria}=    Get From Dictionary    ${current_row}    X

    Safe Input Text    (//input[contains(@id,'textBoxLayerPremiumTRIA')])    ${layer_premium_tria}

    ${commission}=    Get From Dictionary    ${current_row}    Y

    Safe Input Text    (//input[contains(@id,'_textBoxBrokerCommisionPercentage')])    ${commission}

    ${layer_losses_1}=    Get From Dictionary    ${current_row}    Z

    Safe Input Text    (//input[contains(@id,'LayerLosses')])[1]    ${layer_losses_1}

    ${layer_losses_2}=    Get From Dictionary    ${current_row}    AA

    Safe Input Text    (//input[contains(@id,'LayerLosses')])[2]    ${layer_losses_2}

    ${layer_losses_3}=    Get From Dictionary    ${current_row}    AB

    Safe Input Text    (//input[contains(@id,'LayerLosses')])[3]    ${layer_losses_3}

    ${layer_losses_4}=    Get From Dictionary    ${current_row}    AC

    Safe Input Text    (//input[contains(@id,'LayerLosses')])[4]    ${layer_losses_4}

    ${layer_losses_5}=    Get From Dictionary    ${current_row}    AD

    Safe Input Text    (//input[contains(@id,'LayerLosses')])[5]    ${layer_losses_5}

    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Yes'])[4]

    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Yes'])[5]

    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[2]

    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[3]

    Safe Click Element

    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]

    Wait Until Element Is Visible    (//a[normalize-space()='Renewal Comparison'])[1]    20s

    Safe Click Element

    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderCenter$ctl00'])[1]

    Wait Until Element Is Visible    (//a[normalize-space()='Deductible And Coverage'])[1]    20s

    ${policy_form}=    Get From Dictionary    ${current_row}    AE

    Safe Select From List By Label

    ...    (//select[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlDeductibleAndCoverages_ddlPolicyForm'])[1]

    ...    ${policy_form}

    ${peril_insured}=    Get From Dictionary    ${current_row}    AF

    Safe Select From List By Label

    ...    (//select[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlDeductibleAndCoverages_ddlPerilInsured'])[1]

    ...    ${peril_insured}

    ${property_covered}=    Get From Dictionary    ${current_row}    AG

    Safe Select From List By Label

    ...    (//select[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlDeductibleAndCoverages_ddlPropertyCovered'])[1]

    ...    ${property_covered}

    ${comments}=    Get From Dictionary    ${current_row}    AH

    Safe Input Text

    ...    (//textarea[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlDeductibleAndCoverages_textBoxPleaseSpecifyPC'])[1]

    ...    ${comments}

    Safe Click Element

    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]

    Wait Until Element Is Visible    (//a[normalize-space()='Quote Items'])[1]    20s

    ${min_earned_prem}=    Get From Dictionary    ${current_row}    AI

    Safe Input Text

    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlQuoteItems_textBoxMiniumEarnedPremium'])[1]

    ...    ${min_earned_prem}

    ${valuation_property}=    Get From Dictionary    ${current_row}    AJ

    Safe Select From List By Label

    ...    (//select[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlQuoteItems_ddlValuation'])[1]

    ...    ${valuation_property}

    ${business_interuption}=    Get From Dictionary    ${current_row}    AK

    Safe Select From List By Label

    ...    (//select[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlQuoteItems_ddlValuationBusinessInterruption'])[1]

    ...    ${business_interuption}

    ${coinsurance_property}=    Get From Dictionary    ${current_row}    AL

    Safe Input Text

    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlQuoteItems_txtCoinsuranceProperty'])[1]

    ...    ${coinsurance_property}

    ${coinsurance_business}=    Get From Dictionary    ${current_row}    AM

    Safe Input Text

    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlQuoteItems_txtCoinsuranceBusiness'])[1]

    ...    ${coinsurance_business}

    ${perils}=    Get From Dictionary    ${current_row}    AN

    Safe Input Text

    ...    (//textarea[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlQuoteItems_txtPerils'])[1]

    ...    ${perils}

    Safe Click Element

    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]

    Wait Until Element Is Visible    (//a[normalize-space()='Additional Terms And Conditions'])[1]    20s

    ${ocurrence_limit}=    Get From Dictionary    ${current_row}    AO

    Safe Select From List By Label

    ...    (//select[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlTermsAndConditions_ddlOccurrenceLimit'])[1]

    ...    ${ocurrence_limit}

    Safe Click Element

    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]


Fill Pricing Details[Reissue]

    Safe Click Element    (//a[normalize-space()='Contract I + Location'])[1]

    Wait Until Element Is Visible    (//a[normalize-space()='Contract I + Location'])[1]    20s

    Safe Click Element

    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]

    Wait Until Element Is Visible    (//a[normalize-space()='Contract II'])[1]    10s

    Safe Click Element

    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]

    Wait Until Element Is Visible    (//a[normalize-space()='AIR Results Occurence'])[1]    20s

    Safe Click Element

    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_btnNext'])[1]

    Wait Until Element Is Visible    (//a[normalize-space()='Pricing'])[1]    20s

    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Yes'])[1]

    Safe Click Element

    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]

    Wait Until Element Is Visible    (//a[normalize-space()='Renewal Comparison'])[1]    20s

    Safe Click Element

    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderCenter$ctl00'])[1]

    Wait Until Element Is Visible    (//a[normalize-space()='Deductible And Coverage'])[1]    20s

    Safe Click Element

    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]

    Wait Until Element Is Visible    (//a[normalize-space()='Quote Items'])[1]    20s

    Safe Click Element

    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]

    Wait Until Element Is Visible    (//a[normalize-space()='Additional Terms And Conditions'])[1]    20s

    Safe Click Element

    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]


Quote[Copy Risk]

    Wait Until Element Is Visible    (//a[normalize-space()='Finalize'])[1]    20s

    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Yes'])[1]

    Sleep    5s

    Safe Click Element    (//span[normalize-space()='Correct'])[1]

    ${comments_under_quote}=    Get From Dictionary    ${current_row}    AR

    Safe Input Text

    ...    (//textarea[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlFinalize_txtComments'])[1]

    ...    ${comments_under_quote}

    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[3]

    Safe Click Element

    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_btnGenerateQuote'])[1]


Pre Bind Endorsement[Reissue]

    Safe Click Element

    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]

    Wait Until Keyword Succeeds    3x    5s    Wait Until Element Is Visible    (//a[normalize-space()='Subjectivities'])[1]    2s

    Safe Click Element

    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl00'])[1]

    Sleep    5s


Quote[Reissue]

    Wait Until Element Is Visible    (//a[normalize-space()='Finalize'])[1]    20s

    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Yes'])[1]

    Sleep    5s

    Safe Click Element    (//span[normalize-space()='Correct'])[1]

    ${comments_under_quote}=    Get From Dictionary    ${current_row}    AR

    Safe Input Text

    ...    (//textarea[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlFinalize_txtComments'])[1]

    ...    ${comments_under_quote}

    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[3]

    Safe Click Element

    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_btnGenerateQuote'])[1]


Ready To Bind[Reissue]

    Wait Until Element Is Visible    (//a[normalize-space()='Bind'])[1]    30s

    Safe Click Element    (//a[normalize-space()='Bind'])[1]

    Wait Until Element Is Visible    (//legend[@id='fsMainContentLegend'])[1]    30s

    Safe Click Element

    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl01'])[1]

    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]    30s

    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Yes'])[1]

    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[2]

    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[3]
    

    Safe Click Element

    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_btnBind'])[1]


Ready To Bind[Renewal]

    Wait Until Element Is Visible    (//a[normalize-space()='Account Summary'])[1]    30s

    Safe Click Element    (//a[normalize-space()='Account Summary'])[1]

    ${comments_under_bind}=    Get From Dictionary    ${current_row}    AR

    Safe Input Text

    ...    (//textarea[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlAccountSummary_textBoxReferralComments'])[1]

    ...    ${comments_under_bind}

    Safe Select From List By Index

    ...    (//select[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_ctrlAccountSummary_ddlApprover'])[1]

    ...    1

    Click Element    xpath=//body

    Safe Click Element

    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_btnComplete'])[1]

    Sleep    20s

    Safe Click Element    (//a[normalize-space()='Bind'])[1]

    Wait Until Element Is Visible    (//legend[@id='fsMainContentLegend'])[1]    30s

    Safe Click Element

    ...    (//input[@name='ctl00$ctl00$ctl00$PartContentPlaceHolderMain$ContentPlaceHolderMain$ContentPlaceHolderPolicyMain$ctl01'])[1]

    Wait Until Element Is Visible    (//legend[@class='ui-widget ui-widget-header ui-corner-all'])[1]    30s

    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Yes'])[1]

    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[2]

    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='No'])[3]

    Safe Click Element

    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_btnBind'])[1]

    Sleep    90s


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

    ${endorsement_type_list}=    Get From Dictionary    ${current_row}    AW

    Safe Input Text

    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_txtEndorsementType'])[1]

    ...    ${endorsement_type_list}

    Sleep    2s

    Safe Click Element    (//a[normalize-space()='${endorsement_type_list}'])[1]

    Sleep    5s

    Safe Click Element

    ...    (//input[@id='ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_btEditEndorsement'])[1]

    Sleep    3s

    TRY

        ${policy_expiry_date}=    Get Current Date    increment=5 days    result_format=%m/%d/%Y

        Safe Input Text

        ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderEndorsementMain_ContentPlaceHolderEndorsement_DatePickerPolicyExpiryDate_textDate'])[1]

        ...    ${policy_expiry_date}

    EXCEPT

        ${policy_expiry_date}=    Get Current Date    increment=5 days    result_format=%m/%d/%Y

        Safe Input Text

        ...    (//input[@id='ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderEndorsementMain_ContentPlaceHolderEndorsement_DatePickerPolicyExpiryDate_textDate'])[1]

        ...    ${policy_expiry_date}

        Safe Click Element    (//span[normalize-space()='No Premium'])[1]

    END

    Click Element    xpath=//body

    Sleep    5s

    Safe Click Element

    ...    //*[@id="ctl00_ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_endorsementStandardButtons_btnSubmit"]

    Sleep    5s

    Safe Click Element

    ...    //*[@id="ctl00_ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ContentPlaceHolderPolicyMain_BtnBack"]

    Sleep    5s

    Safe Click Element    xpath=//div[@class='RiskHeaderColumnOne']//div[1]


Renewal

    Wait Until Element Is Visible    (//a[normalize-space()='Renew'])[1]    30s

    Safe Click Element    (//a[normalize-space()='Renew'])[1]

    Sleep    10s

    Risk Creation[Renewal]

    Pre Bind Endorsement[Reissue]

    Quote[Reissue]

    Ready To Bind[Renewal]

    Book    False

    Issue Policy


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

    Ready To Bind[Reissue]

    Book    False

    Issue Policy


Copy Risk

    Wait Until Element Is Visible    (//a[normalize-space()='Copy Risk'])[1]    30s

    Safe Click Element    (//a[normalize-space()='Copy Risk'])[1]

    Sleep    10s

    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Yes'])[3]

    Safe Click Element    (//span[@class='ui-button-text'][normalize-space()='Yes'])[4]

    Fill Insured Details[Copy Risk]

    Fill Pricing Details[Copy Risk]

    Pre Bind Endorsement

    Quote[Copy Risk]

    Ready To Bind

    Book    False

    Issue Policy


Continue on Next Risk

    Wait Until Element Is Visible    xpath=//a[normalize-space()='Create New Risk >']    20s

    Safe Click Element    xpath=//a[normalize-space()='Create New Risk >']

    Safe Click Element    xpath=//a[normalize-space()='Create New Risk >']

    Wait Until Element Is Visible    xpath=//a[normalize-space()='US Primary GL']    30s

    Safe Click Element    xpath=//a[normalize-space()='US Primary GL']

    Safe Click Element

    ...    id:ctl00_ctl00_PartContentPlaceHolderMain_ContentPlaceHolderMain_ctrlClearanceSearch_CreateInsured

    Wait Until Element Is Visible    (//span[@class='ui-button-text'][normalize-space()='No'])[1]    40s


Safe Click Element

    [Arguments]    ${locator}

    Wait Until Keyword Succeeds    15x    2s    Click Element When Visible    ${locator}


Safe Input Text

    [Arguments]    ${locator}    ${text}

    Wait Until Keyword Succeeds    15x    2s    Input Text When Element Is Visible    ${locator}    ${text}


Safe Select From List By Label

    [Arguments]    ${locator}    ${label}

    Wait Until Keyword Succeeds    15x    2s    Select From List By Label    ${locator}    ${label}


Safe Select From List By Index

    [Arguments]    ${locator}    ${index}

    Wait Until Keyword Succeeds    15x    2s    Select From List By Index    ${locator}    ${index}


Safe Choose File

    [Arguments]    ${locator}    ${file_path}

    Wait Until Keyword Succeeds    15x    2s    Choose File    ${locator}    ${file_path}
