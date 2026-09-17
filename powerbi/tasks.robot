*** Settings ***
Library     RPA.Browser.Selenium    auto_close=${FALSE}
Library     RPA.Excel.Files
Library     RPA.PDF
Library     RPA.Windows
Library     Collections
Library     RPA.FileSystem
Library     DateTime
Library     String
Library     RPA.Robocorp.WorkItems


*** Variables ***
${DATA_SOURCE}      ${CURDIR}${/}UWP_users.xlsx


*** Tasks ***
PowerBI
    Open Available Browser
    ...    https://app.powerbi.com/groups/me/apps/70bbd38e-c4e1-40ee-88b9-1f8e0d725b1d/reports/aaf71743-5f3e-4341-958c-264e924844db/ReportSectionaab7d51975a936d40c41?experience=power-bi
    ...    maximized=${TRUE}
    ...    browser_selection=edge
    Set Selenium Implicit Wait    10s
    Sleep    5s
    Input Text    (//input[@id='email'])[1]    neha.rana@libertymutual.com
    Click Element When Visible    (//button[normalize-space()='Submit'])[1]
    Sleep    5s
    Input Text    (//input[@id='i0118'])[1]    extra195bonesBRINK2398
    Click Element When Visible    (//input[@id='idSIButton9'])[1]
    Sleep    20s
    Open Workbook    ${DATA_SOURCE}
    ${excel_rows}=    Read Worksheet    Need N numbers
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
        ${current_row}=    Get From List    ${excel_rows}    ${counter-1}
        Log    ${current_row}
        Set Global Variable    ${current_row}
        ${clear_filter}=    Is Element Visible    (//button[@aria-disabled='false'])[1]
        IF    ${clear_filter} == True
            Click Element When Visible    (//button[@aria-disabled='false'])[1]
            Sleep    5s
        END
        Click Element When Visible
        ...    (//div[@class='themeableElement textLabel trimmedTextWithEllipsis'][normalize-space()='Last Name, First Name'])[1]
        Sleep    5s
        ${name}=    Get From Dictionary    ${current_row}    G
        Input Text    (//input[@aria-label='Search'])[1]    ${name}
        Press Keys    (//input[@aria-label='Search'])[1]    RETURN
        Sleep    10s
        ${user_not_exixt}=    Is Element Visible    (//p[@class='slicerEmptyResult'])[1]
        IF    ${user_not_exixt} == True
            Set Cell Value    ${counter}    H    NOT FOUND
            Clear Element Text    (//input[@aria-label='Search'])[1]
            Click Element When Visible
            ...    (//div[@class='themeableElement textLabel trimmedTextWithEllipsis'][normalize-space()='Last Name, First Name'])[1]
        ELSE
            Click Element When Visible    (//span[@class='slicerText'])[1]
            Sleep    10s
            ${entries}=    RPA.Browser.Selenium.Get Text    (//span[@class='slicerCountText'])[1]
            IF    ${entries} > 1
                Click Element When Visible    (//div[contains(text(),'N Number and Status')])[1]
                Sleep    5s
                Input Text    (//input[@placeholder='Search'])[3]    active
                ${user_not_exixt1}=    Is Element Visible    (//p[@class='slicerEmptyResult'])[1]
                IF    ${user_not_exixt1} == True
                    Set Cell Value    ${counter}    H    NOT ACTIVE
                    Click Element When Visible    (//div[contains(text(),'N Number and Status')])[1]
                ELSE
                    Click Element When Visible    (//span[@class='glyphicon checkbox checkboxOutline'])[2]
                    Sleep    10s
                END
            END
            ${n_number}=    RPA.Browser.Selenium.Get Text    (//div[@role='gridcell'])[16]
            ${sbu}=    RPA.Browser.Selenium.Get Text    (//div[@role='gridcell'])[4]
            ${email}=    RPA.Browser.Selenium.Get Text    (//div[@role='gridcell'])[12]
            ${user_status}=    RPA.Browser.Selenium.Get Text    (//div[@role='gridcell'])[14]
            Log    ${counter}
            Set Cell Value    ${counter}    I    ${n_number}
            Set Cell Value    ${counter}    H    ${entries}
            Set Cell Value    ${counter}    K    ${sbu}
            Set Cell Value    ${counter}    J    ${user_status}
            Set Cell Value    ${counter}    L    ${email}
            Click Element When Visible
            ...    (//div[@class='themeableElement textLabel trimmedTextWithEllipsis'][normalize-space()='Last Name, First Name'])[1]
            Click Element When Visible    (//div[contains(text(),'N Number and Status')])[1]
            Save Workbook
        END
    END


*** Keywords ***
Flow to execute
    Search and copy values in excel

Login to UI
    Open Available Browser
    ...    https://app.powerbi.com/groups/me/apps/70bbd38e-c4e1-40ee-88b9-1f8e0d725b1d/reports/aaf71743-5f3e-4341-958c-264e924844db/ReportSectionaab7d51975a936d40c41?experience=power-bi
    ...    maximized=${TRUE}
    ...    browser_selection=edge
    Set Selenium Implicit Wait    10s
    Sleep    5s
    Input Text    (//input[@id='email'])[1]    neha.rana@libertymutual.com
    Click Element When Visible    (//button[normalize-space()='Submit'])[1]
    Sleep    5s
    Input Text    (//input[@id='i0118'])[1]    extra195bonesBRINK2398
    Click Element When Visible    (//input[@id='idSIButton9'])[1]
    Sleep    20s

Search and copy values in excel
    ${clear_filter}=    Is Element Visible    (//button[@aria-disabled='false'])[1]
    IF    ${clear_filter} == True
        Click Element When Visible    (//button[@aria-disabled='false'])[1]
    END
    Click Element When Visible
    ...    (//div[@class='themeableElement textLabel trimmedTextWithEllipsis'][normalize-space()='Last Name, First Name'])[1]
    Sleep    5s
    ${name}=    Get From Dictionary    ${current_row}    G
    Input Text    (//input[@aria-label='Search'])[1]    ${name}
    Press Keys    (//input[@aria-label='Search'])[1]    RETURN
    Sleep    10s
    Click Element When Visible    (//span[@class='slicerText'])[1]
    Sleep    10s
    ${text}=    RPA.Browser.Selenium.Get Text    (//span[@class='slicerCountText'])[1]
    ${n_number}=    RPA.Browser.Selenium.Get Text    (//div[@role='gridcell'])[16]
    ${sbu}=    RPA.Browser.Selenium.Get Text    (//div[@role='grid'])[2]
    ${sbu}=    RPA.Browser.Selenium.Get Text    (//div[@role='gridcell'])[12]
    ${user_status}=    RPA.Browser.Selenium.Get Text    (//div[@role='gridcell'])[14]
    Set Cell Value    ${counter}    I    ${n_number}
