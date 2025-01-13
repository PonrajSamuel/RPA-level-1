*** Settings ***
Documentation       Insert the sales data for the week and export it as a PDF.
Library    RPA.Browser.Selenium    auto_close=${FALSE}
Library    RPA.HTTP
Library    RPA.Excel.Files
Library    RPA.Windows
Library    RPA.PDF
Library    OperatingSystem
*** Variables ***
${OUTPUT_DIR}        C:\\Users\\Ponraj\\OneDrive - Yitro Business Consultants India Private Limited\\Pictures\\Saved Pictures
*** Tasks ***
Insert the sales data for the week and export it as a PDF
  Open the intranet website
  Log in
  Download the Excel file
  # Fill and submit the form fr one person
  Fill the form using the data from the Excel file
  Collect the results
  Export the table as a PDF
  Log out and close the browser
*** Keywords ***
Open the intranet website
    Open Available Browser   https://robotsparebinindustries.com/
    Maximize Browser Window
    Sleep    5s
Log in    
  Input Text    //*[@id="username"]    maria
  Input Password    //*[@id="password"]    thoushallnotpass
  Submit Form
  Wait Until Page Contains Element    id:sales-form
  Log    Done.
  Sleep    5s
Download the Excel file  
  Download    https://robotsparebinindustries.com/SalesData.xlsx    overwrite=True

Fill and submit the form for one person 
  [Arguments]    ${sales_rep}
  Input Text    //*[@id="firstname"]    ${sales_rep}[First Name]
  Input Text    //*[@id="lastname"]    ${sales_rep}[First Name]
  Input Text    //*[@id="salesresult"]   ${sales_rep}[Sales]
  Select From List By Value     //*[@id="salestarget"]      ${sales_rep}[Sales Target]
  Click Button    //*[@id="sales-form"]/button
Fill the form using the data from the Excel file
  Open Workbook    SalesData.xlsx
  ${sales_reps}=    Read Worksheet As Table    header=True
  Close Workbook
  FOR    ${sales_rep}    IN    @{sales_reps}
        Fill and submit the form for one person    ${sales_rep}
  END
Collect the results
    RPA.Browser.Selenium.Screenshot    //*[@id="root"]/div/div/div/div[2]/div[1]    ${OUTPUT_DIR}${/}sales_summary.png
Export the table as a PDF    
    Wait Until Element Is Visible    //*[@id="sales-results"]/table
    ${sales_results_html}=    Get Element Attribute    //*[@id="sales-results"]/table    outerHTML
    Html To Pdf    ${sales_results_html}    ${OUTPUT_DIR}${/}sales_results.pdf
    Sleep    5s
Log out and close the browser
    Click Button    //*[@id="logout"]    
    Close Browser
    