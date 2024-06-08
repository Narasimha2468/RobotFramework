# Keyword Classification
# F = Framework, W = Web, M = Mobile, B = Both

*** Settings ***
Library     ExcellentLibrary
Library     Collections
Library     OperatingSystem
Library     JSONLibrary
Library     String
Library     DateTime
Library     Utilities.py
# Library    Selenium2Library
Library     DatabaseLibrary
Library     SSHLibrary
Library		IBrogEmailSmtp.py
# Library    AppiumLibrary
# Suite Setup    Set Excel File


*** Variables ***
${IBorgTemplatePath}         Test_Data (IDOCX).xlsx
${UserProfile}    True
${CurrentTestCaseId}
${TempDict}
&{TempReportData}
${SkipTeardown}    True
${PreviousTestCase}   DummyValue
${RecordScreen}    NO
${ScreenRecordingStarted}    False
${ReportGenerated}    False
${EmailSent}    False
${Attachments}    None
${EmailTotalTestCasesCount}    0
${EmailFailedTestCaseCount}    0
${EmailPassedTestCaseCount}    0
${EmailSkippedTestCaseCount}    0
${EmailTotalScriptFailCount}    0
${EmailTotalAssertionFailCount}    0
${EmailTotalFailCount}    0
${EmailDuration}    0
@{ExecuteTestCaseList}
@{ReportData}
@{StepWiseReportData}
@{PassedTestCases}
@{FailedTestCases}
@{SkippedTestCases}
@{CurrentTestCaseAssertionFails}
@{TotalAssertionFails}
@{TotalScriptFails}
@{AllFailedTestCaseInfo}
${HistoricalReportpath}    ${CURDIR}\\Templates\\Historical_Report.xlsx

&{DictAlphaNum}		'0'=7		'1'=8		'2'=9		'3'=10		'4'=11		'5'=12		'6'=13		'7'=14		'8'=15		'9'=16		'A'=29		'B'=30		'C'=31		'D'=32		'E'=33		'F'=34		'G'=35		'H'=36		'I'=37		'J'=38		'K'=39		'L'=40		'M'=41		'N'=42		'O'=43		'P'=44		'Q'=45		'R'=46		'S'=47		'T'=48		'U'=49		'V'=50		'W'=51		'X'=52		'Y'=53		'Z'=54
&{DictSpecButtons}    ENTER=66    SPACE=62    TAB=61    SHIFT=59    ALT=57    CTRL=113    BACK=4
&{DictSpecChar}    '!'=33    '@'=64    '#'=35    '$'=36    '%'=37    '^'=94    '&'=38    '*'=42    '('=40    ')'=41
&{DatabaseModules}     PostgreSQL=psycopg2    Oracle=cx_Oracle     MySQL=pymysql    MicrosoftSQLServer=pyodbc


@{0_arglist}        RELOAD      UNSELECT IFRAME     BROWSER CLOSE    CLOSE APP    NO STEP      GO BACK     READ SERVER LOG    GET SYSTEM INFORMATION    GET ALL RUNNING PROCESSES    SWIPE UP
@{1_arglist}        READ SERVER LOG UNTIL       WRITE TO SERVER     NAVIGATE TO URL    NAVIGATE TO MOBILE URL     CLICK       SELECT IFRAME    ELEMENT SHOULD BE PRESENT    ELEMENT SHOULD NOT BE PRESENT    DOUBLE CLICK    CHECK ELEMENT IS ENABLED    CHECK ELEMENT IS DISABLED    CLEAR TEXT FIELD    SET DOWNLOAD DIRECTORY    HANDLE ALERT BOX    PRESS BUTTON    TRANSFER FILE TO DEVICE    OPEN APP    CONNECT TO DB    EXECUTE SQL QUERY    FETCH DATA FROM SQL DB    RIGHT CLICK    TABLE SHOULD BE PRESENT    DELETE ALL ROWS FROM THE TABLE   SCROLL ELEMENT TO VIEW      GET THE TEXT    CREATE EXCEL FILE    WRITE TO EXCEL FILE    READ FROM EXCEL FILE    CREATE A DIRECTORY    DELETE THE DIRECTORY    EMPTY THE DIRECTORY    RENAME THE FILE    RENAME THE DIRECTORY    UNZIP FILE    CONVERT XLS FILE TO XLSX    CONVERT PDF FILE TO DOCX    CONVERT DOCX FILE TO PDF    MOVE A FILE    MOVE A DIRECTORY    COPY A FILE    COPY A DIRECTORY    CAPTURE SCREENSHOT    CALCULATE FILE OR DIRECTORY SIZE    WRITE TO A FILE    APPEND TO A FILE    READ A FILE    PERFORM MATH OPERATION    TERMINATE A PROCESS    CHECK PROCESS STATUS    GET PROCESS INFORMATION    EXECUTE PYTHON CODE    EXECUTE JAVASCRIPT CODE    ENCRYPT DATA    DECRYPT DATA    SEND EMAIL WITH ATTACHMENTS    SEND EMAIL WITHOUT ATTACHMENTS    COMPRESS FOLDER    MERGE EXCEL FILES    GENERATE FAKE TEST DATA    VERIFY FILE EXISTS    DELETE A FILE

@{2_arglist}        LOGIN TO SERVER     OPEN SERVER CONNECTION      INPUT DATA       UPLOAD FILE    WAIT FOR ELEMENT    CLICK LOOP    WAIT FOR ELEMENT TO BE GONE    STORE ELEMENT VALUE    SELECT FROM DROPDOWN
@{3_arglist}
@{Verifylist}        VERIFY TEXT    VERIFY THE TEXT FIELD VALUE    VERIFY THE TEXT FIELD VALUE CONTAINS    VERIFY ELEMENT COUNT    VERIFY TEXT CONTAINS    VERIFY TEXT STARTS WITH    VERIFY TEXT ENDS WITH
@{Exceptionlist1}    VERIFY PAGE TITLE
@{Exceptionlist2}    VERIFY ATTRIBUTE VALUE
@{FunctionKeywords}     IF_ELEMENTVISIBLE      ELSE     ELIF       FOR_ELEMENTVISIBLE      TRY     EXCEPT     END


*** Keywords ***
# To execute keywords having zero arguments (F)
Keyword_arg0
    [Arguments]        ${Keyword}
    Run Keyword     ${Keyword}


# To execute keywords having one argument (F)
Keyword_arg1
    [Arguments]        ${Keyword}        ${arg1}        ${arg2}
    ${value}        Run Keyword     ${Keyword}       ${arg1}        ${arg2}
    [Return]        ${value}


# To execute keywords having two arguments (F)
Keyword_arg2
    [Arguments]        ${Keyword}        ${arg1}        ${arg2}     ${arg3}
    Run Keyword     ${Keyword}       ${arg1}        ${arg2}     ${arg3}


# To execute keywords having three arguments (F)
Keyword_arg3
    [Arguments]        ${Keyword}        ${arg1}        ${arg2}     ${arg3}
    Run Keyword     ${Keyword}       ${arg1}        ${arg2}     ${arg3}

# To execute keywords having four arguments (F)
Keyword_arg4
    [Arguments]        ${Keyword}        ${arg1}        ${arg2}     ${arg3}    ${arg4}
    Run Keyword     ${Keyword}       ${arg1}        ${arg2}     ${arg3}    ${arg4}

# To execute keywords having one argument and return a value (F)
FunctionKeyword_arg1
    [Arguments]        ${Keyword}        ${arg1}        ${arg2}
    ${value}        Run Keyword     ${Keyword}       ${arg1}        ${arg2}
    [Return]      ${value}

# To execute keywords having two argument and return a value (F)
FunctionKeyword_arg2
    [Arguments]        ${Keyword}        ${arg1}        ${arg2}     ${arg3}
    ${value}        Run Keyword     ${Keyword}       ${arg1}        ${arg2}     ${arg3}
    [Return]      ${value}

# To send the report (F)
Send Email
    [Arguments]    ${Config}    ${Email}    ${Body}    ${Subject}    ${Attachments}
    IF    "${Attachments}" != "None"
        Ibrog HTMLEmail AttachY    ${Config}    iborg.automation@sirmaindia.com    ${Email}    ${Attachments}    ${Subject}    Hello,<br><br>${Body}<br><br>For any queries please call back to us on Mob: +91 986754321 or Email id: iborg.automation@sirmaindia.com<br><br>Thanks,<br>Automation Team
    ELSE
        Ibrog_HTMLEmail_AttachN    ${Config}    iborg.automation@sirmaindia.com    ${Email}    ${Subject}    Hello,<br><br>${Body}<br><br>For any queries please call back to us on Mob: +91 986754321 or Email id: iborg.automation@sirmaindia.com<br><br>Thanks,<br>Automation Team
    END
    Set Global Variable    ${EmailSent}    True
    Log to file    Email sent to '${Email}'

# To find and set the excel file (F)
Set Excel File
    ${IBorgTemplatePath}    Fetch Excel File
    IF    len(${IBorgTemplatePath}) != 0
        Set Global Variable    ${IBorgTemplatePath}    ${IBorgTemplatePath}[0]
    ELSE
        FAIL    The Excel file named Test_Data could not be located in the downloads folder.
    END

# To navigate chrome browser to the given URL (W)
Open Chrome Browser to Page
    [Documentation]     Opens Google Chrome to a given web page.
    [Arguments]    ${URL}    ${Element}
    Close All Browsers
    # to kill all before automation chrome instance
    # Evaluate		Utilities.kill_excel_process()
    ${DwnldStat}        Run Keyword and return status       Variable Should exist       ${DownloadPath}
	${DownloadPath}     Set Variable If     ${DwnldStat}            ${DownloadPath}         %{USERPROFILE}\Downloads
	${chrome_options}=    Evaluate    sys.modules['selenium.webdriver'].ChromeOptions()    sys
    IF  "${Browser}".upper().strip() == "HEADLESS CHROME"
        Call Method    ${chrome_options}    add_argument    --headless
        Call Method    ${chrome_options}    add_argument    --disable-gpu
    END
    Call Method    ${chrome_options}    add_argument    --disable-extensions
    Call Method    ${chrome_options}    add_argument    --disable-web-security
    Call Method    ${chrome_options}    add_argument    --allow-running-insecure-content
    Call Method    ${chrome_options}    add_argument    --safebrowsing-disable-extension-blacklist
    Call Method    ${chrome_options}    add_argument    --safebrowsing-enable
    Call Method    ${chrome_options}    add_argument    --ignore-certificate-errors
    Call Method    ${chrome_options}    add_argument    --disable-extensions
    Call Method    ${chrome_options}    add_argument    --disable-infobars
    Call Method    ${chrome_options}    add_argument    --safebrowsing-disable-download-protection
    Call Method    ${chrome_options}    add_argument    --start-maximized
    # Uncomment the below line if profile level configuration is required
    # Run keyword If     ${UserProfile}     Call Method     ${chrome_options}     add_argument     --user-data-dir\=%{USERPROFILE}/AppData/Local/Google/Chrome/User Data/Profile 2
    ${excludeopts}=     Create List     enable-automation   load-extension
    Call Method  ${chrome_options}  add_experimental_option  excludeSwitches   ${excludeopts}
    ${prefs}=  Create Dictionary  credentials_enable_service  ${False}
    Set To Dictionary    ${prefs}       profile.default_content_settings.popups  1
    Set To Dictionary       ${prefs}         download.default_directory=${DownloadPath}
    Call Method    ${chrome_options}    add_experimental_option     prefs   ${prefs}
    Create Webdriver    Chrome    IBorgChrome     options=${chrome_options}
    Go To   ${URL}

    FOR    ${i}    IN RANGE    5
        ${ActualURL}    Get Location
        IF    "${ActualURL}" in "${URL}"
            Set Global Variable    ${Iborg_SYS_BrwOpened}    ${True}
            Log to file     Chrome is launched to URL ${URL}
            Exit For Loop
        ELSE IF     "${URL}" in "${ActualURL}"
            Set Global Variable    ${Iborg_SYS_BrwOpened}    ${True}
            Log to file     Chrome is launched to URL ${URL}
            Exit For Loop
        ELSE
            Reload Page
            Sleep    3s
        END
    END


    # ${SelInstance}	get library instance		Selenium2Library
	# ${browserName}	Evaluate		$SelInstance.driver.capabilities['browserName']
	# ${browserVersion}	Evaluate		$SelInstance.driver.capabilities['browserVersion']


# To navigate MS Edge browser to the given URL (W)
Open MS Edge Browser to Page
	[Documentation]    Opens Microsoft Egde to a given web page.
	[Arguments]    ${URL}    ${Element}
	Close All Browsers
    ${options}=    Evaluate    sys.modules['selenium.webdriver'].EdgeOptions()    sys, selenium.webdriver
    IF  "${Browser}".upper().strip() == "HEADLESS MS EDGE"
        Call Method    ${options}    add_argument    --headless
        Call Method    ${options}    add_argument    --disable-gpu
    END
    Call Method	    ${options}	  add_argument	  --disable-extensions
	Call Method	    ${options}	  add_argument	  --disable-web-security
	Call Method	    ${options}	  add_argument	  --allow-running-insecure-content
	Call Method	    ${options}	  add_argument	  --safebrowsing-disable-extension-blacklist
	Call Method    	${options}	  add_argument	  --safebrowsing-enable
	Call Method	    ${options}	  add_argument	  --ignore-certificate-errors
	Call Method	    ${options}	  add_argument	  --disable-extensions
	Call Method	    ${options}	  add_argument	  --no-sandbox
	Call Method	    ${options}	  add_argument	  --disable-infobars
	Call Method	    ${options}	  add_argument	  --safebrowsing-disable-download-protection
	Call Method	    ${options}	  add_argument	  --start-maximized
    Create Webdriver    Edge        options=${options}
    Go To    ${URL}

    FOR    ${i}    IN RANGE    5
        ${ActualURL}    Get Location
        IF    "${ActualURL}" in "${URL}"
            Set Global Variable    ${Iborg_SYS_BrwOpened}    ${True}
            Log to file     MS Edge is launched to URL ${URL}
            Exit For Loop
        ELSE
            Reload Page
            Sleep    3s
        END
    END


# To navigate Firefox browser to the given URL (W)
Open Firefox Browser to Page
	[Documentation]	Opens Mozilla Firefox to a given web page.
	[Arguments]    ${URL}    ${Element}
	Close All Browsers
    ${options}=    Evaluate    sys.modules['selenium.webdriver'].firefox.options.Options()    sys, selenium.webdriver.firefox.options
    IF  "${Browser}".upper().strip() == "HEADLESS FIREFOX"
        Call Method    ${options}    add_argument    --headless
        Call Method    ${options}    add_argument    --disable-gpu
    END
    Call Method	    ${options}	  add_argument	  --disable-extensions
	Call Method	    ${options}	  add_argument	  --disable-web-security
	Call Method	    ${options}	  add_argument	  --allow-running-insecure-content
	Call Method	    ${options}	  add_argument	  --safebrowsing-disable-extension-blacklist
	Call Method    	${options}	  add_argument	  --safebrowsing-enable
	Call Method	    ${options}	  add_argument	  --ignore-certificate-errors
	Call Method	    ${options}	  add_argument	  --disable-extensions
	Call Method	    ${options}	  add_argument	  --no-sandbox
	Call Method	    ${options}	  add_argument	  --disable-infobars
	Call Method	    ${options}	  add_argument	  --safebrowsing-disable-download-protection
	Call Method	    ${options}	  add_argument	  --start-maximized
	# ${fp}=	Evaluate	sys.modules['selenium.webdriver'].FirefoxProfile()	sys
	# Call Method    	${fp}	set_preference	browser.download.folderList	2
	# Call Method	    ${fp}	set_preference	browser.download.dir	os.getcwd()
	# Call Method	    ${fp}	set_preference	browser.helperApps.neverAsk.saveToDisk	text/pdf, application/pdf, text/plain, application/vnd.ms-excel, text/csv, text/comma-separated-values, application/octet-stream, text/csv,application/x-msexcel,application/excel,application/x-excel,application/vnd.ms-excel,image/png,image/jpeg,text/html,text/plain,application/msword,application/xml, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
	# Call Method	    ${fp}	set_preference	browser.helperApps.neverAsk.openFile	text/pdf, application/pdf, text/plain, application/vnd.ms-excel, text/csv, text/comma-separated-values, application/octet-stream, text/csv,application/x-msexcel,application/excel,application/x-excel,application/vnd.ms-excel,image/png,image/jpeg,text/html,text/plain,application/msword,application/xml, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
	# Call Method	    ${fp}	set_preference	browser.download.manager.showWhenStarting	false
	# Call Method	    ${fp}	set_preference	browser.download.dir	**path**
	# Call Method	    ${fp}	set_preference	browser.helperApps.alwaysAsk.force	false
	# Call Method	    ${fp}	set_preference	browser.download.manager.alertOnEXEOpen	false
	# Call Method	    ${fp}	set_preference	browser.download.manager.focusWhenStarting	false
	# Call Method	    ${fp}	set_preference	browser.download.manager.useWindow	false
	# Call Method	    ${fp}	set_preference	browser.download.manager.showAlertOnComplete	false
	# Call Method	    ${fp}	set_preference	browser.download.manager.closeWhenDone	false
	# Create Webdriver    Firefox    profile=${fp}
	Create Webdriver    Firefox    options=${options}
	Maximize Browser Window
	Go To	${URL}

    FOR    ${i}    IN RANGE    5
        ${ActualURL}    Get Location
        IF    "${ActualURL}" in "${URL}"
            Set Global Variable    ${Iborg_SYS_BrwOpened}    ${True}
            Log to file     Firefox is launched to URL ${URL}
            Exit For Loop
        ELSE
            Reload Page
            Sleep    3s
        END
    END

# To launch browser based on user's choice
Navigate to url
	[Arguments]    ${URL}    ${Element}
	Run Keyword If	"${Browser}".upper().strip() in ['CHROME','HEADLESS CHROME']	Open Chrome Browser to Page    ${URL}    ${Element}
	...	ELSE IF	"${Browser}".upper().strip() in ['MS EDGE','HEADLESS MS EDGE']	Open MS Edge Browser to Page	${URL}    ${Element}
	...	ELSE IF	"${Browser}".upper().strip() in ['FIREFOX','HEADLESS FIREFOX']	Open Firefox Browser to Page	${URL}    ${Element}
	...	ELSE	Open Chrome Browser to Page    ${URL}    ${Element}


# To remove temporary files which are created during runtime (F)
Remove Temp Files
    [Arguments]    ${Directories}
    ${DirList}    Evaluate    "${Directories}".split(",")
    FOR    ${dir}    IN    @{DirList}
        ${DirExist}    Run keyword and ignore error    OperatingSystem.Directory Should Exist    ${CURDIR}\\${dir}
        IF    "${DirExist}[0]" == "PASS"
            ${Files}    OperatingSystem.List files in directory    ${CURDIR}\\${dir}    absolute=true
            FOR    ${file}    IN    @{Files}
                Remove file    ${file}
            END
        END
    END


# To create the log file and directories (F)
Initial Test Setup
    [Arguments]    ${Group}
    ${SuiteData}    Load Json From File    ${CURDIR}\\Test Suite Info\\TestSuiteInfo.json
    Set Global Variable    ${TimeStamp}    ${SuiteData}[report_dir_time_stamp]
    Set Global Variable    ${TestSuiteStartTime}    ${SuiteData}[suite_start_time]
    Setup Logger    ${ReportPath}\\Execution Summary\\Execution_Summary_${TimeStamp}\\Group_${Group}\\Log\\LogFile_${TimeStamp}.log
    Create Directory    ${ReportPath}\\Execution Summary\\Execution_Summary_${TimeStamp}\\Group_${Group}\\Snapshots
    # Create Directory    ${ReportPath}\\Execution Summary\\Execution_Summary_${TimeStamp}\\Group_${Group}\\Videos
    Create Directory    ${ReportPath}\\Execution Summary\\Execution_Summary_${TimeStamp}\\Group_${Group}\\Reports


# To generate and format the report (F)
Teardown
    ${TestSuiteEndTime}    Get Current Date     result_format=%H:%M:%S

    IF     "${SkipTeardown}" == "False" and len(${TempReportData}) !=0
        FOR     ${ReportDict}    IN    @{ReportData}
            ${DictKeys}    Get Dictionary Keys    ${ReportDict}    sort_keys=False
            FOR     ${keys}    IN    @{DictKeys}
                IF     """${ReportDict['${keys}']}""" == """PASSED"""
                    Append to list    ${PassedTestCases}    ${ReportDict['TestCaseID']}
                ELSE IF     """${ReportDict['${keys}']}""" == """FAILED"""
                    Append to list    ${FailedTestCases}    ${ReportDict['TestCaseID']}
                ELSE IF    """${ReportDict['${keys}']}""" == """SKIPPED"""
                    Append to list    ${SkippedTestCases}    ${ReportDict['TestCaseID']}
                END
            END
        END

        ${TotalTestCasesCount}    Evaluate    len(${ExecuteTestCaseList})
        ${FailedTestCaseCount}    Evaluate    len(${FailedTestCases})
        ${PassedTestCaseCount}    Evaluate    len(${PassedTestCases})
        ${SkippedTestCaseCount}    Evaluate    len(${SkippedTestCases})
        ${TotalScriptFailCount}    Evaluate    len(${TotalScriptFails})
        ${TotalAssertionFailCount}    Evaluate    len(${TotalAssertionFails})
        ${TotalFailCount}    Evaluate    ${TotalScriptFailCount}+${TotalAssertionFailCount}
        ${Duration}   Subtract Time From Time    ${TestSuiteEndTime}    ${TestSuiteStartTime}

        ${AllBrowserExecutionCompleted}    ${BrowserCount}    Check if execution completed
        IF      ${AllBrowserExecutionCompleted} and "${RecordScreen}".upper().strip() == "YES"
            Create File         ${CURDIR}\\VideoRec_Info\\stop_recording.txt
            Wait Until Created        ${CURDIR}\\VideoRec_Info\\testrecord.mp4
            sleep        1s
            Move File        ${CURDIR}\\VideoRec_Info\\testrecord.mp4        ${ReportPath}\\Execution Summary\\Execution_Summary_${TimeStamp}\\screenrecord.mp4
        END

        ${FailTestCaseInfo}    Group Failed Testcases    ${ReportData}
        Add To Json    group_${Group}    total_testcases_count    ${TotalTestCasesCount}
        Add To Json    group_${Group}    failed_testcase_count    ${FailedTestCaseCount}
        Add To Json    group_${Group}    passed_testcase_count    ${PassedTestCaseCount}
        Add To Json    group_${Group}    skipped_testcase_count    ${SkippedTestCaseCount}
        Add To Json    group_${Group}    total_scriptfail_count    ${TotalScriptFailCount}
        Add To Json    group_${Group}    total_assertionfail_count    ${TotalAssertionFailCount}
        Add To Json    group_${Group}    total_fail_count    ${TotalFailCount}
        Add To Json    group_${Group}    duration    ${Duration}
        Add To Json    group_${Group}    failed_testcase_info    ${FailTestCaseInfo}

        IF    "${GenerateReport.upper().strip()}" == "YES"
            Generate Report    ${StepWiseReportData}    StepWiseSummary
            Generate Report    ${ReportData}    TestCaseSummary
            Format Report    ${ReportPath}\\Execution Summary\\Execution_Summary_${TimeStamp}\\Group_${Group}\\Reports\\TestCase_Report_${TimeStamp}.xlsx    TestCaseSummary
            Format Report    ${ReportPath}\\Execution Summary\\Execution_Summary_${TimeStamp}\\Group_${Group}\\Reports\\StepWise_Report_${TimeStamp}.xlsx    StepWiseSummary
            ${ReportGenerated}    Set Variable    True
        END

        IF     "${SendEmail.upper().strip()}" == "YES" and ${AllBrowserExecutionCompleted}
            # ${Duration}   Subtract Time From Time    ${TestSuiteEndTime}    ${TestSuiteStartTime}
            ${Duration}    Convert To Minutes Or Seconds    ${Duration}
            IF     ${ReportGenerated}
                ${Result}    Compress Execution Summary With 7zip    ${ReportPath}\\Execution Summary    ${TimeStamp}
                ${Attachments}    Set Variable If    ${Result}    ${ReportPath}\\Execution Summary\\Execution_Summary_${TimeStamp}\\Compressed Archive\\Execution_Summary_${TimeStamp}.7z
            END

            ${SuiteData}    Load Json From File    ${CURDIR}\\Test Suite Info\\TestSuiteInfo.json
            FOR    ${group}    IN    group_g1    group_g2    group_g3    group_g4    group_g5
                ${EmailTotalTestCasesCount}    Evaluate    ${EmailTotalTestCasesCount} + ${SuiteData}[${group}][total_testcases_count]
                ${EmailFailedTestCaseCount}    Evaluate    ${EmailFailedTestCaseCount} + ${SuiteData}[${group}][failed_testcase_count]
                ${EmailPassedTestCaseCount}    Evaluate    ${EmailPassedTestCaseCount} + ${SuiteData}[${group}][passed_testcase_count]
                ${EmailSkippedTestCaseCount}    Evaluate    ${EmailSkippedTestCaseCount} + ${SuiteData}[${group}][skipped_testcase_count]
                ${EmailTotalScriptFailCount}   Evaluate    ${EmailTotalScriptFailCount} + ${SuiteData}[${group}][total_scriptfail_count]
                ${EmailTotalAssertionFailCount}    Evaluate    ${EmailTotalAssertionFailCount} + ${SuiteData}[${group}][total_assertionfail_count]
                ${EmailTotalFailCount}    Evaluate    ${EmailTotalFailCount} + ${SuiteData}[${group}][total_fail_count]
                ${EmailDuration}    Evaluate    ${EmailDuration} + ${SuiteData}[${group}][duration]
                IF    isinstance(${SuiteData}[${group}][failed_testcase_info], list)
                    Run Keyword If    len(${SuiteData}[${group}][failed_testcase_info])!=0     Append To List    ${AllFailedTestCaseInfo}    ${SuiteData}[${group}][failed_testcase_info]
                END
            END

            ${EmailDuration}    Convert To Minutes Or Seconds    ${EmailDuration}
            ${Subject}    Set Variable    Regression Test Result Summary for the '${ApplicationName}' build '${ApplicationVersion}'
            ${Body}    Set Variable    Please find the below test execution result summary for the '${ApplicationName}' build '${ApplicationVersion}'.<br><br>Test cases execution time : '${EmailDuration}'.<br><br>Test cases execution platform : '${OperatingSystem}'.<br><br>

            IF    "${AutomationType}" != "Mobile"
                ${Body}    Catenate    ${Body}    Test cases execution browser count : '${BrowserCount}'.<br><br>
            END

            ${SummaryTable}    Set Variable    <b>Test Suite execution summary ::</b><br><br><table cellspacing="0" cellpadding="7" border="1" style="text-align:center"><thead bgcolor="82E0AA"><th>Sr. No</th><th>Passed</th><th>Failed</th><th>Skipped</th><th>Total</th></thead><tbody><tr><td>1</td><td>${EmailPassedTestCaseCount}</td><td>${EmailFailedTestCaseCount}</td><td>${EmailSkippedTestCaseCount}</td><td>${EmailTotalTestCasesCount}</td></tr></tbody></table><br><br><b>Failed test cases summary ::</b><br><br><table cellspacing="0" cellpadding="7" border="1" style="text-align:center"><thead bgcolor="85C1E9"><th>Sr. No</th><th>Script failed</th><th>Assertion failed</th><th>Total</th></thead><tbody><tr><td>1</td><td>${EmailTotalScriptFailCount}</td><td>${EmailTotalAssertionFailCount}</td><td>${EmailTotalFailCount}</td></tr></tbody></table>
            ${Body}    Catenate    ${Body}    ${SummaryTable}

            IF    len(${AllFailedTestCaseInfo}) != 0
                ${FailedTestCaseSummaryTable}    Set Variable    <br><br><table cellspacing="0" cellpadding="7" border="1" style="text-align:center"><thead bgcolor="fff280"><th>Sr. No</th><th>Failed Test case ID</th><th>Test Case Description</th><th>Failed Reason</th></thead><tbody>
                ${SrNum}    Set Variable    ${0}
                FOR    ${TestCases}    IN    @{AllFailedTestCaseInfo}
                    FOR    ${Data}    IN    @{TestCases}
                        ${SrNum}    Evaluate    ${SrNum}+1
                        ${FailedTestCaseSummaryTable}    Catenate    ${FailedTestCaseSummaryTable}    <tr><td>${SrNum}</td><td>${Data}[test_case_id]</td><td>${Data}[test_case_desc]</td><td>${Data}[reason]</td></tr>
                    END
                END
                ${FailedTestCaseSummaryTable}    Catenate    ${FailedTestCaseSummaryTable}    </tbody></table>
                ${Body}    Catenate    ${Body}    ${FailedTestCaseSummaryTable}
            END

            Send Email     ${CURDIR}\\IBorgSmtpConfig-iborg.automation.json    ${Email}    ${Body}    ${Subject}    ${Attachments}
        END


        # Update historical report
        IF    "${Regressiontest.upper().strip()}" == "YES" and ${AllBrowserExecutionCompleted}
            ${PassPercent}    Evaluate    round((int(${PassedTestCaseCount})/int(${TotalTestCasesCount}))*100,2)
            ${Date}    Get Current Date    result_format=%d-%m-%Y
            ${TestSuiteStartTime}    Convert To 12hr Format    ${TestSuiteStartTime}
            ${TestSuiteEndTime}    Convert To 12hr Format    ${TestSuiteEndTime}
            IF    not ${EmailSent}
                ${Duration}    Convert To Minutes Or Seconds    ${Duration}
            END

            ${MetaData}    Create Dictionary    SerialNum=${EMPTY}    Date=${Date}    Project=${ApplicationName}    Build=${ApplicationVersion}    Platform=${AutomationType}    ExecutionStartTime=${TestSuiteStartTime}    ExecutionEndTime=${TestSuiteEndTime}    Duration=${Duration}    TotalPassed=${PassedTestCaseCount}    TotalFailed=${FailedTestCaseCount}    TotalSkipped=${SkippedTestCaseCount}    GrandTotal=${TotalTestCasesCount}    PassPercentage=${PassPercent}%

            Open workbook    ${HistoricalReportpath}
            ${SheetData}    Read sheet data
            ${SheetRows}    Get Length    ${SheetData}
            ${cellNum}    Evaluate    ${SheetRows}+1

            Set To Dictionary    ${MetaData}    SerialNum    ${SheetRows}
            ${DictKeys}    Get Dictionary Keys    ${MetaData}    sort_keys=False

            FOR     ${col}    ${keys}    IN ENUMERATE    @{DictKeys}
                ${cell}    Evaluate    chr(65+${col})
                Write To Cell    ${cell}${cellNum}    ${MetaData['${keys}']}
            END
            Save
            Close Workbook
            Cell Border    Sheet    ${1}    HistoricalReport    ${HistoricalReportpath}
        END
    END

# To generate the final report (F)
Generate Report
    [Arguments]    ${ReportData}    ${ReportName}
    IF    "${ReportName}" == "TestCaseSummary"
        Create Workbook    ${ReportPath}\\Execution Summary\\Execution_Summary_${TimeStamp}\\Group_${Group}\\Reports\\TestCase_Report_${TimeStamp}.xlsx
        Write to cell    D9    Execution Start Time
        Write to cell    E9    Duration
        Write to cell    F9    Status
        Write to cell    G9    Error Type
        Write to cell    H9    Reason
    ELSE
        Create Workbook    ${ReportPath}\\Execution Summary\\Execution_Summary_${TimeStamp}\\Group_${Group}\\Reports\\StepWise_Report_${TimeStamp}.xlsx
        Write to cell    D9    Steps
        Write to cell    E9    Step Name
        Write to cell    F9    Execution Start Time
        Write to cell    G9    Step Duration
        Write to cell    H9    Status
        Write to cell    I9    Error Type
        Write to cell    J9    Reason
    END

    Write to cell    A1    Application Name
    Write to cell    A2    Application Version
    Write to cell    A3    Operating System
    Write to cell    A4    Environment
    Write to cell    B1    ${ApplicationName}
    Write to cell    B2    ${ApplicationVersion}
    Write to cell    B3    ${Operating System}
    Write to cell    B4    ${Environment}

    IF     "${AutomationType}" == "Web"
        Write to cell    A5    Browser
        Write to cell    A6    URL
        Write to cell    B5    ${Browser}
        Write to cell    B6    ${URL}
    END

    Write to cell    A9    Serial No.
    Write to cell    B9    Test Case ID
    Write to cell    C9    Test Case Description


    FOR     ${row}    ${ReportDict}    IN ENUMERATE    @{ReportData}
        ${ExecutionTime}    calculate_execution_time    ${ReportPath}\\Execution Summary\\Execution_Summary_${TimeStamp}\\Group_${Group}\\Log\\LogFile_${TimeStamp}.log
        ${ActiveTestcase}    Set Variable     ${ReportDict['TestCaseID']}

        IF    "${ReportName}" == "TestCaseSummary"
            Set To Dictionary    ${ReportDict}    ExecutionStartTime     ${ExecutionTime['${ReportDict['TestCaseID']}']['start_time']}
            Set To Dictionary    ${ReportDict}    Duration     ${ExecutionTime['${ReportDict['TestCaseID']}']['elapsed_time']}
        END

        ${DictKeys}    Get Dictionary Keys    ${ReportDict}    sort_keys=False

        # To set step numbers in the report
        IF    "${ActiveTestCase}" != "${PreviousTestCase}"
            Set Global Variable    ${SerialNum}    0
        END
        ${SerialNum}    Evaluate    ${SerialNum}+1
        Set to Dictionary    ${ReportDict}    StepNum    Step-${SerialNum}

        FOR     ${col}    ${keys}    IN ENUMERATE    @{DictKeys}
            ${cell}    Evaluate    chr(65+${col})
            ${CellNum}    Evaluate    ${row}+10
            Write to Cell    ${cell}${CellNum}    ${ReportDict['${keys}']}
        END
        Set Global Variable    ${PreviousTestCase}    ${ActiveTestCase}
    END
    Save
    Close Workbook


# To set log file (F)
Setup Logger
    [Arguments]    ${loggerfilepath}
    Create File    ${loggerfilepath}
    Set Global Variable    ${loggerfilepath}


# To write the logs (F)
Log to file
    [Arguments]    ${loggingtext}    ${loglevel}=INFO
    # ${loggingtext}    Replace String    ${loggingtext}    \n    ${SPACE}
    # ${loggingtext}    Evaluate    "${loggingtext}".strip()
    ${loglevel}    Evaluate    "${loglevel}".strip().upper()
    ${dateTime}    Get Current Date
    Append To File     ${loggerfilepath}      [${dateTime}] - [${loglevel}] - ${loggingtext}\n


# To set the download directory before launching the browser (W)
Set download directory
    [Arguments]    ${Directorypath}    ${Element}
    OperatingSystem.Directory Should Exist    ${Directorypath}
    ${Directorypath}   Normalize path    ${Directorypath}
    Set Global Variable    ${DownloadPath}    ${Directorypath}
    Log to file    Download directory set to path '${Directorypath}'


#Connect to Database (B)
Connect to DB
    [Arguments]    ${TestData}    ${Element}
    ${Data}    	Split String	${TestData}	    ,
    ${JSONData}    Utilities.Convert String To Json    ${Data}
    Check and Install Module    ${DatabaseModules}[${JSONData}[dbmanagementsystem]]
    Connect To Database    ${DatabaseModules}[${JSONData}[dbmanagementsystem]]    ${JSONData}[dbname]    ${JSONData}[dbuser]    ${JSONData}[dbpass]    ${JSONData}[dbhost]    ${JSONData}[dbport]
    Log to file    Successfully Connected to '${JSONData}[dbname]' Database


#Execute SQL query (B)
Execute SQL Query
    [Arguments]    ${TestData}    ${Element}
    ${isQueryFile}    Evaluate    "${TestData}".endswith(".sql")
    IF    ${isQueryFile}
        Execute SQL Script    ${TestData}
    ELSE
        Execute SQL String    ${TestData}
    END
    Log to file    Query '${TestData}' executed successfully


# To fetch the data from SQL DB (B)
Fetch Data From SQL DB
    [Arguments]    ${TestData}    ${Element}
    ${QueryResults}    Query    ${TestData}    returnAsDict=True
    log     ${QueryResults}    warn
    Add To Json    query_results    result    ${QueryResults}


# To delete all rows from a table (B)
Delete All Rows From The Table
    [Arguments]    ${TestData}    ${Element}
    Delete All Rows From Table    ${TestData}
    Log to file    Successfully deleted all rows of table '${TestData}'


# To verify if the given table is present or not (B)
Table should be present
    [Arguments]    ${TestData}    ${Element}
    ${Result}    Run keyword and return status     Table Must Exist    ${TestData}
    IF    ${Result}
        Log to file    '${TestData}' table exists
    ELSE
        Set Global Variable    ${AssertionError}    True
		Append to list     ${CurrentTestCaseAssertionFails}    '${TestData}' table does not exist
        Fail    '${TestData}' table does not exist
    END


# To verify if the text contains the given text (B)
Verify text contains
    [Arguments]    ${Xpath}    ${ExpectedResult}    ${Element}
	FOR		${i}	IN RANGE	5
		${Status}	Run keyword and return status 	Wait until element is visible     ${Xpath}
		IF 	${Status}
			${ActualResult}    Get Text    ${Xpath}
            ${IsVariable}    ${doesntContainNegation}    ${VariableName}    Check if its a variable    ${ExpectedResult}

            IF    ${IsVariable} and ${doesntContainNegation}
                ${SuiteData}    Load Json From File    ${CURDIR}\\Test Suite Info\\TestSuiteInfo.json
                ${ExpectedResult}    Set Variable    ${SuiteData}[variables][${VariableName}]
            END
			${AssertionStatus}        Evaluate    "${ExpectedResult}" in "${ActualResult}"
			IF         ${AssertionStatus}
				Log        Check point got passed since excepcted : '${ExpectedResult}' text is present in actual : '${ActualResult}' text        WARN
				Log to file     Check point got passed since expected: '${ExpectedResult}' text is present in actual: '${ActualResult}' text
			ELSE
				Set Global Variable    ${AssertionError}    True
				Append to list     ${CurrentTestCaseAssertionFails}    Check point got failed since excepcted : '${ExpectedResult}' text is not present in actual : '${ActualResult}' text
				Fail    Check point got failed since expected : '${ExpectedResult}' text is not present in actual : '${ActualResult}' text
			END
			Exit For Loop
		ELSE
			IF	"${AutomationType}" == "Mobile"
                # ${xpath}    Get WebElement    ${xpath}
				Swipe By Percent		50    40    50   1	500
				Sleep	1s
			ELSE
				Exit For Loop
			END
		END
        # Run Keyword If     ${i}==4    Fail    Element with locator ${Xpath} not visibile/found    #under test
	END


# To verify if the text starts with the given text (B)
Verify text starts with
    [Arguments]    ${Xpath}    ${ExpectedResult}    ${Element}
    FOR		${i}	IN RANGE	5
	    ${Status}	Run keyword and return status 	Wait until element is visible     ${Xpath}
		IF 	${Status}
            ${ActualResult}    Get Text    ${Xpath}
            ${IsVariable}    ${doesntContainNegation}    ${VariableName}    Check if its a variable    ${ExpectedResult}

            IF    ${IsVariable} and ${doesntContainNegation}
                ${SuiteData}    Load Json From File    ${CURDIR}\\Test Suite Info\\TestSuiteInfo.json
                ${ExpectedResult}    Set Variable    ${SuiteData}[variables][${VariableName}]
            END
            ${AssertionStatus}        Evaluate    "${ActualResult}".startswith("${ExpectedResult}")
            IF         ${AssertionStatus}
                Log        Check point got passed since '${ActualResult}' starts with '${ExpectedResult}' text        WARN
                Log to file     Check point got passed since '${ActualResult}' starts with '${ExpectedResult}' text
            ELSE
                Set Global Variable    ${AssertionError}    True
                Append to list     ${CurrentTestCaseAssertionFails}    Check point got failed since '${ActualResult}' doesn not start with '${ExpectedResult}' text
                Fail    Check point got failed since '${ActualResult}' doesn not start with '${ExpectedResult}' text
            END
        	Exit For Loop
		ELSE
			IF	"${AutomationType}" == "Mobile"
                # ${xpath}    Get WebElement    ${xpath}
				Swipe By Percent		50    40    50   1	500
				Sleep	1s
			ELSE
				Exit For Loop
			END
		END
        # Run Keyword If     ${i}==4    Fail    Element with locator ${Xpath} not visibile/found    #under test
	END


# To verify if the text ends with the given text (B)
Verify text ends with
    [Arguments]    ${Xpath}    ${ExpectedResult}    ${Element}
    FOR		${i}	IN RANGE	5
	    ${Status}	Run keyword and return status 	Wait until element is visible     ${Xpath}
		IF 	${Status}
            ${ActualResult}    Get Text    ${Xpath}
            ${IsVariable}    ${doesntContainNegation}    ${VariableName}    Check if its a variable    ${ExpectedResult}

            IF    ${IsVariable} and ${doesntContainNegation}
                ${SuiteData}    Load Json From File    ${CURDIR}\\Test Suite Info\\TestSuiteInfo.json
                ${ExpectedResult}    Set Variable    ${SuiteData}[variables][${VariableName}]
            END
            ${AssertionStatus}        Evaluate    "${ActualResult}".endswith("${ExpectedResult}")
            IF         ${AssertionStatus}
                Log        Check point got passed since '${ActualResult}' ends with '${ExpectedResult}' text        WARN
                Log to file     Check point got passed since '${ActualResult}' ends with '${ExpectedResult}' text
            ELSE
                Set Global Variable    ${AssertionError}    True
                Append to list     ${CurrentTestCaseAssertionFails}    Check point got failed since '${ActualResult}' doesn not end with '${ExpectedResult}' text
                Fail    Check point got failed since '${ActualResult}' doesn not end with '${ExpectedResult}' text
            END
            Exit For Loop
		ELSE
			IF	"${AutomationType}" == "Mobile"
                # ${xpath}    Get WebElement    ${xpath}
				Swipe By Percent		50    40    50   1	500
				Sleep	1s
			ELSE
				Exit For Loop
			END
		END
        # Run Keyword If     ${i}==4    Fail    Element with locator ${Xpath} not visibile/found    #under test
	END

# To verify attribute value (W)
Verify attribute value
    [Arguments]    ${Xpath}    ${TestData}    ${ExpectedResult}    ${Element}
    Wait until element is visible     ${Xpath}    20s
    ${ActualResult}    Get Element Attribute    ${Xpath}   ${TestData}
    ${IsVariable}    ${doesntContainNegation}    ${VariableName}    Check if its a variable    ${ExpectedResult}

    IF    ${IsVariable} and ${doesntContainNegation}
        ${SuiteData}    Load Json From File    ${CURDIR}\\Test Suite Info\\TestSuiteInfo.json
        ${ExpectedResult}    Set Variable    ${SuiteData}[variables][${VariableName}]
    END
    ${AssertionStatus}        Set Variable if      "${ActualResult}"=="${ExpectedResult}"        True        False

    IF     ${AssertionStatus}
        Log        Check point got passed since expected value: ${ExpectedResult} and actual value: ${ActualResult} are matched.        WARN
        Log to file     Check point got passed since expected value: '${ExpectedResult}' and actual value: '${ActualResult}' are matched.
    ELSE
        Set Global Variable    ${AssertionError}    True
        Append to list     ${CurrentTestCaseAssertionFails}    Check point got failed since excepcted value : '${ExpectedResult}' and actual value : '${ActualResult}' are not matched
        Fail    Check point got failed since expected value: '${ExpectedResult}' and actual value: '${ActualResult}' are not matched.
    END


# To verify the element count (B)
Verify element count
    [Arguments]    ${Xpath}    ${ExpectedResult}    ${Element}
    IF    "${AutomationType}" == "Mobile"
        # ${xpath}    Get WebElement    ${xpath}
        ${ActualResult}    Get Matching Xpath Count    ${Xpath}
    ELSE
        Wait until element is visible        ${Xpath}    20s
        ${ActualResult}    Get element count    ${Xpath}
    END
    ${IsVariable}    ${doesntContainNegation}    ${VariableName}    Check if its a variable    ${ExpectedResult}

    IF    ${IsVariable} and ${doesntContainNegation}
        ${SuiteData}    Load Json From File    ${CURDIR}\\Test Suite Info\\TestSuiteInfo.json
        ${ExpectedResult}    Set Variable    ${SuiteData}[variables][${VariableName}]
    END
    ${AssertionStatus}        Set Variable if      "${ActualResult}"=="${ExpectedResult}"        True        False

    IF     ${AssertionStatus}
        Log        Check point got passed since expected: ${ExpectedResult} and actual: ${ActualResult} counts are matched.        WARN
        Log to file     Check point got passed since expected: '${ExpectedResult}' and actual: '${ActualResult}' counts are matched.
    ELSE
        Set Global Variable    ${AssertionError}    True
        Append to list     ${CurrentTestCaseAssertionFails}    Check point got failed since excepcted : '${ExpectedResult}' and actual : '${ActualResult}' counts are not matched
        Fail    Check point got failed since expected: '${ExpectedResult}' and actual: '${ActualResult}' counts are not matched.
    END


# To verify the value present in text field (B)
Verify the text field value
    [Arguments]    ${Xpath}   ${ExpectedResult}    ${Element}
    IF    "${AutomationType}" == "Mobile"
        # ${xpath}    Get WebElement    ${xpath}
        FOR		${i}	IN RANGE	5
            ${Status}	Run keyword and return status 	Wait until element is visible     ${Xpath}
            IF    ${Status}
                ${ActualResult}    Get text    ${Xpath}
                Exit For Loop
            ELSE
                Swipe By Percent		50    40    50   1	500
				Sleep	1s
            END
        END
        Run Keyword If     ${i}==4    Fail    Element with locator ${Xpath} not visibile/found
    ELSE
        Click    ${Xpath}    ${Element}
        Evaluate    pyperclip.copy('')
        Press Keys    NONE    CTRL+A
        Press Keys    NONE    CTRL+C

        IF    "${ExpectedResult}" == "None"
            ${ExpectedResult}   Set Variable    ${EMPTY}
        END
        ${ActualResult}    Evaluate    pyperclip.paste()
    END

    ${IsVariable}    ${doesntContainNegation}    ${VariableName}    Check if its a variable    ${ExpectedResult}

    IF    ${IsVariable} and ${doesntContainNegation}
        ${SuiteData}    Load Json From File    ${CURDIR}\\Test Suite Info\\TestSuiteInfo.json
        ${ExpectedResult}    Set Variable    ${SuiteData}[variables][${VariableName}]
    END

    ${AssertionStatus}        Set Variable if      "${ActualResult}"=="${ExpectedResult}"        True        False
    IF         ${AssertionStatus}
        Log        Check point got passed since expected: ${ExpectedResult} and actual: ${ActualResult} values are matched.        WARN
        Log to file     Check point got passed since expected: '${ExpectedResult}' and actual: '${ActualResult}' values are matched.
    ELSE
        Set Global Variable    ${AssertionError}    True
        Append to list     ${CurrentTestCaseAssertionFails}    Check point got failed since excepcted : '${ExpectedResult}' and actual : '${ActualResult}' values are not matched
        Fail    Check point got failed since expected: '${ExpectedResult}' and actual: '${ActualResult}' values are not matched.
    END


# To verify the value present in text field (B)
Verify the text field value contains
    [Arguments]    ${Xpath}   ${ExpectedResult}    ${Element}
    IF    "${AutomationType}" == "Mobile"
        ${xpath}    Get WebElement    ${xpath}
        FOR		${i}	IN RANGE	5
            ${Status}	Run keyword and return status 	Wait until element is visible     ${Xpath}
            IF    ${Status}
                ${ActualResult}    Get text    ${Xpath}
                Exit For Loop
            ELSE
                Swipe By Percent		50    40    50   1	500
				Sleep	1s
            END
        END
    ELSE
        Click    ${Xpath}    ${Element}
        Evaluate    pyperclip.copy('')
        Press Keys    NONE    CTRL+A
        Press Keys    NONE    CTRL+C

        IF    "${ExpectedResult}" == "None"
            ${ExpectedResult}   Set Variable    ${EMPTY}
        END
        ${ActualResult}    Evaluate    pyperclip.paste()
        ${ActualResult}    Evaluate    "${ActualResult}".split('.')[0]

    END

    ${IsVariable}    ${doesntContainNegation}    ${VariableName}    Check if its a variable    ${ExpectedResult}

    IF    ${IsVariable} and ${doesntContainNegation}
        ${SuiteData}    Load Json From File    ${CURDIR}\\Test Suite Info\\TestSuiteInfo.json
        ${ExpectedResult}    Set Variable    ${SuiteData}[variables][${VariableName}]
    END

    ${AssertionStatus}        Set Variable if      "${ActualResult}" in "${ExpectedResult}"        True        False
    IF         ${AssertionStatus}
        Log        Check point got passed since expected: ${ExpectedResult} and actual: ${ActualResult} values are matched.        WARN
        Log to file     Check point got passed since expected: '${ExpectedResult}' and actual: '${ActualResult}' values are matched.
    ELSE
        Set Global Variable    ${AssertionError}    True
        Append to list     ${CurrentTestCaseAssertionFails}    Check point got failed since excepcted : '${ExpectedResult}' and actual : '${ActualResult}' values are not matched
        Fail    Check point got failed since expected: '${ExpectedResult}' and actual: '${ActualResult}' values are not matched.
    END

# To click on the specified element (B)
Click
    [Arguments]    ${Xpath}    ${Element}
    IF      "${AutomationType}"=="Mobile"
        # ${xpath}    Get WebElement    ${xpath}
        FOR     ${iter}      IN RANGE        10
            # ${xpath}    Get WebElement    ${xpath}
            ${Visible}       Run Keyword and Return Status       Wait Until Element is visible      ${xpath}
            IF       "${Visible}"=="True"
                AppiumLibrary.Click Element        ${xpath}
                Exit for loop
            ELSE
                Swipe By Percent		50    45    50    1    500
                # Swipe    500    500    500    1
                Sleep	1s
            END
        END
        Run keyword If    ${iter}==9    Fail    Element with locator '${Xpath}' not visible/found
    ELSE
        # ${scrollable}     Run keyword and return status   Scroll Element into View    ${Xpath}
        # ${logicallypresent}     Run keyword and return status   Scroll Element into View    ${Xpath}
        Wait until element is visible     ${Xpath}    20s
        Scroll Element into View    ${Xpath}
        Click Element     ${Xpath}
    END
    Log to file     Clicked on the '${Element}' button


# Select from dropdownlist (W)
Select From Dropdown
    [Arguments]    ${Xpath}    ${Value}    ${Element}
    Wait Until Element is Visible    ${Xpath}    20S
    Select From List By Label    ${Xpath}    ${Value}
    Log to file    Selected '${Value}' from dropdown '${Element}'


# To right click on the specified element (W)
Right Click
    [Arguments]    ${Xpath}    ${Element}
    Wait until element is visible     ${Xpath}    20s
    Scroll Element into View    ${Xpath}
    Open Context Menu     ${Xpath}
    Log to file     Right clicked on '${Element}' button

# To double-click on the specified element (W)
Double Click
    [Arguments]    ${Xpath}    ${Element}
    Wait until element is visible     ${Xpath}    20s
    Scroll Element into View    ${Xpath}
    Double Click Element     ${Xpath}
    Log to file     Clicked on '${Element}' button

# To wait for the element to be gone (B)
Wait for element to be gone
    [Arguments]    ${Xpath}    ${TimeOut}    ${Element}
    Run keyword and return status    Wait Until Page Does Not Contain Element    ${Xpath}    ${TImeOut}
    Log to file    Waited for '${Element}' to be not visible

# To press the keyboard keys (B)
Press Button
    [Arguments]    ${Data}    ${Element}
    IF    "${AutomationType}" == "Mobile"
        ${DictKeys}    Get Dictionary Keys    ${DictSpecButtons}    sort_keys=False
        IF    "${Data}" in "${DictKeys}"
            Press Keycode    ${DictSpecButtons}[${Data}]
        ELSE
            ${Keys}    Evaluate    list("${Data}")

            FOR     ${Key}    IN    @{Keys}
                IF    "${Key}" == "@"
                    # Press Keycode    62
                    Press Keycode    ${DictAlphaNum}['2']    metastate=1
                ELSE
                    ${IsUpper}    Evaluate    "${Key}".isupper()
                    ${Key}    Evaluate    "${Key}".upper()
                    IF    ${IsUpper}
                        Press Keycode    ${DictAlphaNum}['${Key}']    metastate=1
                    ELSE
                        Press Keycode    ${DictAlphaNum}['${Key}']
                    END
                END
            END
        END
    ELSE
        Press keys    NONE    ${Data}
    END
    Log to file    Pressed '${Data}'


# To input text on the specified field (B)
Input Data
    [Arguments]    ${Xpath}    ${Data}    ${Element}
    IF      "${AutomationType}"=="Mobile"
        # Log    ${Xpath}    error
        # ${xpath}    Get WebElement    ${xpath}
        FOR     ${iter}      IN RANGE        10
            ${Visible}       Run Keyword and Return Status       Wait Until Page Contains Element      ${xpath}
            IF       "${Visible}"=="True"

                Input Text        ${xpath}        ${Data}
                Exit for loop
            ELSE
                Sleep	1s
                # Swipe By Percent		50    40    50    10
                Swipe    500    500    500    1
            END
        END
        Run keyword If    ${iter}==9    Fail    Element with locator '${Xpath}' not visible/found
    ELSE
        Wait until element is visible     ${Xpath}    20s
        Scroll Element into View    ${Xpath}
        Input text     ${Xpath}    ${Data}
    END

    Log to file     Inputed text '${Data}' in the '${Element}' field


# To clear a text field (B)
Clear text field
    [Arguments]    ${Xpath}    ${Element}
    IF	"${AutomationType}" == "Web"
	Wait until element is visible     ${Xpath}    20s
	Scroll Element into View    ${Xpath}
	Click    ${Xpath}    ${Element}
	Press Keys    NONE    CTRL+A
	Press Keys    NONE    BACKSPACE
    ELSE
	Clear Text	${Xpath}
    END
    Log to file     Cleared text in the '${Element}' field


Check if its a variable
    [Arguments]    ${ExpectedResult}
    ${IsVariable}    Evaluate	"""${ExpectedResult}""".startswith("var $")
    ${doesntContainNegation}    Evaluate    not """${ExpectedResult}""".startswith("!")

    ${Match}    Set Variable    None
    IF    ${IsVariable} and ${doesntContainNegation}
        ${Match}=    Evaluate	re.search('\\$(\\w+)', "${ExpectedResult}").group(1)
    END
    [Return]    ${IsVariable}    ${doesntContainNegation}    ${Match}


# To verify if a given text is present or not (B)
Verify text
    [Arguments]    ${Xpath}    ${ExpectedResult}    ${Element}
    FOR     ${i}    IN RANGE    5
        ${Status}    Run keyword and return status    Wait until Element Is Visible           ${Xpath}
        IF    ${Status}
            ${ActualResult}        Get Text        ${Xpath}
            ${IsVariable}    ${doesntContainNegation}    ${VariableName}    Check if its a variable    ${ExpectedResult}

            IF    ${IsVariable} and ${doesntContainNegation}
                ${SuiteData}    Load Json From File    ${CURDIR}\\Test Suite Info\\TestSuiteInfo.json
                ${ExpectedResult}    Set Variable    ${SuiteData}[variables][${VariableName}]
            END

            ${AssertionStatus}        Set Variable if      "${ActualResult}"=="${ExpectedResult}"        True        False
            IF         ${AssertionStatus}
                Log        Check point got passed since excepcted : ${ExpectedResult} and actual : ${ActualResult} values are matched.
                Log to file     Check point got passed since excepcted : '${ExpectedResult}' and actual : '${ActualResult}' values are matched.
            ELSE
                Set Global Variable    ${AssertionError}    True
                Append to list     ${CurrentTestCaseAssertionFails}    Check point got failed since excepcted : '${ExpectedResult}' and actual : '${ActualResult}' values are not matched
                Fail    Check point got failed since excepcted : '${ExpectedResult}' and actual : '${ActualResult}' values are not matched.
            END
            Exit For Loop
        ELSE
            IF	"${AutomationType}" == "Mobile"
                # ${xpath}    Get WebElement    ${xpath}
                Swipe By Percent		50    40    50   1	500
                Sleep	1s
            ELSE
                Exit For Loop
            END
        END
        # Run Keyword If     ${i}==4    Fail    Element with locator ${Xpath} not visibile/found    #under test
	END


# To verify page title (W)
Verify page title
    [Arguments]    ${ExpectedResult}    ${Element}
    ${ActualResult}        Get Title
    ${IsVariable}    ${doesntContainNegation}    ${VariableName}    Check if its a variable    ${ExpectedResult}

    IF    ${IsVariable} and ${doesntContainNegation}
        ${SuiteData}    Load Json From File    ${CURDIR}\\Test Suite Info\\TestSuiteInfo.json
        ${ExpectedResult}    Set Variable    ${SuiteData}[variables][${VariableName}]
    END

    ${AssertionStatus}        Set Variable if      "${ActualResult}"=="${ExpectedResult}"        True        False
    IF         ${AssertionStatus}
        Log        Check point got passed since excepcted : ${ExpectedResult} and actual : ${ActualResult} page titles are matched.        WARN
        Log to file     Check point got passed since excepcted : '${ExpectedResult}' and actual : '${ActualResult}' page titles are matched.
    ELSE
        Set Global Variable    ${AssertionError}    True
        Append to list     ${CurrentTestCaseAssertionFails}   Check point got failed since excepcted : '${ExpectedResult}' and actual : '${ActualResult}' page titles are not matched.
        Fail    Check point got failed since excepcted : '${ExpectedResult}' and actual : '${ActualResult}' page titles are not matched.
    END


# To verify files exists or not
Verify File Exists
    [Arguments]    ${TestData}    ${Element}
    ${FilePath}    Normalize Path    ${TestData}
    ${Status}    Run keyword and Return Status    OperatingSystem.File Should Exist    ${FilePath}
    IF    ${Status}
        Log to file    ${FilePath} file exists
    ELSE
        Set Global Variable    ${AssertionError}    True
        Append to list     ${CurrentTestCaseAssertionFails}   File doesnt exists.
        Fail     File doesnt exists.
    END

# To delete the given file
Delete a file
    [Arguments]    ${TestData}    ${Element}
    ${FilePath}    Normalize Path    ${TestData}
    ${Status}    Run keyword and Return Status    OperatingSystem.File Should Exist    ${FilePath}
    IF    ${Status}
        Remove File    ${FilePath}
        Log to file      ${Filepath} file is deleted
    ELSE
        Fail     Unable to delete, File doesnt exists.
    END



# To handel alerts (W)
Handle alert box
    [Arguments]    ${TestData}    ${Element}
    ${TestData}    Evaluate    "${TestData}".upper()
    Handle alert    ${TestData}    10s
    Log to file    Alert was '${TestData}'

# To close the running mobile app (M)
Close App
    Close Application
    Log to file    Application is closed

# To launch the mobile app (M)
Open App
    [Arguments]    ${TestData}    ${Element}
    Close All Applications
    ${Data}    	Split String	${TestData}	    ,
    ${JSONData}    Utilities.Convert String To Json    ${Data}    False
    ${UpdatedJSONData}    ${RemovedKeys}    Utilities.Remove Keys From Json    json_data=${JSONData}    keys_to_remove=remoteurl,alias
    ${Capabilities}		Create Dictionary

	FOR    ${i}    IN    @{UpdatedJSONData}
		${CapPairs}	Split String	${i}	=
		${Key}	Set variable	${CapPairs}[0]
		${Value}	Set variable	${CapPairs}[1]
		Set To Dictionary	${Capabilities}    ${Key}=${Value}
	END
    Open Application		remote_url=${RemovedKeys}[remoteurl]    alias=Mobile   &{Capabilities}
    #noReset=true    newCommandTimeout=600    asyncTrace=true    shouldUseCompactResponses=true    skipServerInstallation=true    skipDeviceInitialization=true
    Log to file    Application is launched with capabilities "${Capabilities}"


# To go to given url on android or ios device (M)
Navigate to Mobile URL
    [Arguments]    ${TestData}    ${Element}
    Go to URL    ${TestData}
    Log to file    Navigated to '${TestData}'

# To transafer a file from system to device (M)
Transfer File To Device
    [Arguments]    ${Data}    ${Element}
    ${Data}    	Split String	${Data}	    ,
    ${JSONData}    Utilities.Convert String To Json    ${Data}
    Transfer files    ${JSONData}[frompath]    ${JSONData}[topath]    ${JSONData}[adbpath]
    Log to file    Transfered file from '${JSONData}[frompath]' to '${JSONData}[topath]'


# To verify if a given element is present or not (B)
Element Should Be Present
    [Arguments]    ${Xpath}    ${Element}
    Wait Until Element Is Visible    ${Xpath}    20s
    ${Status}    Run Keyword And Return Status    Element Should Be Visible    ${Xpath}
    IF     ${Status}
        Log to file     '${Element}' is visible on screen
    ELSE
        Set Global Variable    ${AssertionError}    True
        Append to list     ${CurrentTestCaseAssertionFails}    ${Element} is not visible on screen
        Fail    '${Element}' is not visible on screen
    END


# To verify if a given element is present or not (B)
Element Should Not Be Present
    [Arguments]    ${Xpath}    ${Element}
    ${Status}    Run Keyword And Return Status    Element Should Be Visible    ${Xpath}
    # expecting element not to be visible
    IF     ${Status}
        Set Global Variable    ${AssertionError}    True
        Append to list     ${CurrentTestCaseAssertionFails}    ${Element} is visible on screen
        Fail    '${Element}' is visible on screen
    ELSE
        Log to file     '${Element}' is not visible on screen
    END


# To verify if a given element is enabled (B)
Check Element is enabled
    [Arguments]    ${Xpath}    ${Element}
    Wait Until Element Is Visible    ${Xpath}    20s
    ${Status}    Run Keyword And Return Status    Element Should Be Enabled    ${Xpath}
    IF     ${Status}
        Log to file     '${Element}' is enabled
    ELSE
        Set Global Variable    ${AssertionError}    True
        Append to list     ${CurrentTestCaseAssertionFails}    ${Element} is not enabled
        Fail    '${Element}' is not enabled
    END


# To verify if a given element is disabled (W)
Check Element is disabled
    [Arguments]    ${Xpath}    ${Element}
    Wait Until Element Is Visible    ${Xpath}    20s
    ${Status}    Run Keyword And Return Status    Element Should Be Disabled    ${Xpath}
    IF     ${Status}
        Log to file     '${Element}' is disabled
    ELSE
        Set Global Variable    ${AssertionError}    True
        Append to list     ${CurrentTestCaseAssertionFails}    ${Element} is not disabled on screen
        Fail    '${Element}' is not disabled
    END


# To select the specified iframe (W)
Select Iframe
    [Arguments]    ${Xpath}    ${Element}
    Wait until element is visible     ${Xpath}
    Select Frame     ${Xpath}
    Log to file    Selected Frame ${Element}


# To wait for a given element to be visible on the screen (B)
Wait for Element
    [Arguments]    ${Xpath}    ${time}      ${Element}
    Wait until element is visible       ${Xpath}    ${time}
    Log to file     Waited for the '${Element}' to be visible


# To wait until the element is enabled (W)
Wait until enabled
    [Arguments]    ${Xpath}    ${Testdata}    ${Element}
    Wait Until Element Is Visible    ${Xpath}    20s
    ${Status}    Run Keyword And Return Status    Wait Until Element Is Enabled    ${Xpath}    ${TestData}
    IF     ${Status}
        Log to file     Waited for the '${Element}' to be enabled
    ELSE
        Set Global Variable    ${AssertionError}    True
        Append to list     ${CurrentTestCaseAssertionFails}    ${Element} is not enabled after ${TestData}s
        Fail    '${Element}' is not enabled after '${TestData}'s
    END


# To unselect all previously selected iframes (W)
Unselect Iframe
    Unselect Frame
    Log to file     Unselected all frames

Get The Text
    [Arguments]    ${Xpath}    ${Element}
    Wait until element is visible     ${Xpath}    20s
    # Scroll Element into View    ${Xpath}
    ${text}     Get Text         ${Xpath}
    log     ${text}     WARN
    Log to file     Text on the '${Element}' button is ${text}
    [Return]        ${text}

Go Back
    Go Back
    Log to file     Navigated Back on Site

# To reload the current page (W)
Reload
    Reload Page
    Log to file     Page reloaded


# To close browser (W)
Browser close
    Close Browser
    Log to file    Browser closed

No Step
    No Operation
    Log to file    Performed 'No operation'

# To upload the given file (W)
# Upload File
#     [Arguments]    ${Xpath}    ${FilePath}    ${Element}
#     Wait Until Element Is Visible    ${Xpath}    20s
#     Choose File    ${Xpath}    ${FilePath}
#     Log to file     Uploaded file '${FilePath}' to ${Element}

Upload File
    [Arguments]    ${Xpath}    ${FilePath}    ${Element}
    ${FilePath}    Split String    ${FilePath}    ,
    IF        len(${FilePath})==1
        Choose File        ${Xpath}       ${FilePath[0]}
    ELSE
        Choose File        ${Xpath}        ${FilePath[0]}\n${FilePath[1]}
    END

    Log to file     Uploaded file '${FilePath}' to ${Element}


# To click on an element given number of times (B)
Click Loop
    [Arguments]    ${Xpath}    ${Value}    ${Element}
    FOR     ${i}      IN RANGE      1   ${Value}+1
        Click        ${Xpath}    ${Element}
    END
    Log to file     Clicked on '${Element}' button '${Value}' times


# Scrolls the web element into view (W)
Scroll Element to View
    [Arguments]    ${Xpath}    ${Element}
    Wait until element is visible     ${Xpath}
    Scroll Element Into View     ${Xpath}
    Log to file    Scrolled Element Into View ${Element}


# Swips up the mobile screen (M)
Swipe Up
    IF    "${AutomationType}" == "Mobile"
        Swipe By Percent		50    40    50   1	500
        Log to file    Swipped up
    END

# Creates a new directoty at given path
Create a Directory
    [Arguments]    ${TestData}    ${Element}
    ${DirPath}    Normalize Path    ${TestData}
    Create Directory    ${DirPath}
    Log to file    Directory created at location '${DirPath}'


# Deletes the directoty at given path
Delete the Directory
    [Arguments]    ${TestData}    ${Element}
    ${DirPath}    Normalize Path    ${TestData}
    Remove Directory    ${DirPath}    recursive=True
    Log to file    Deleted the directory located at '${DirPath}'


# Empties the directoty at given path
Empty the Directory
    [Arguments]    ${TestData}    ${Element}
    ${DirPath}    Normalize Path    ${TestData}
    Empty Directory    ${DirPath}
    Log to file    Emptied the directory located at '${DirPath}'


# Renames the file
Rename the file
    [Arguments]    ${TestData}    ${Element}
    ${FilePath}    Normalize Path    ${TestData}
    Move File    ${FilePath}    ${FilePath}
    Log to file    File renamed to '${FilePath}'


# Renames the directory
Rename the Directory
    [Arguments]    ${TestData}    ${Element}
    ${DirPath}    Normalize Path    ${TestData}
    Move Directory    ${DirPath}    ${DirPath}
    Log to file    Directory renamed to '${DirPath}'


# Unzips the zipped file
Unzip File
    [Arguments]    ${TestData}    ${Element}
    ${Data}    	Split String	${TestData}	    ,
    ${JSONData}    Utilities.Convert String To Json    ${Data}
    Unzip the given file    ${JSONData}[filepath]    ${JSONData}[outputpath]    ${JSONData}[password]
    Log to file    Successfully extracted '${JSONData}[filepath]' to '${JSONData}[outputpath]'


Convert XLS file to XLSX
    [Arguments]    ${TestData}    ${Element}
    ${Data}    	Split String	${TestData}	    ,
    ${JSONData}    Utilities.Convert String To Json    ${Data}
    Convert XLS to XLSX    ${JSONData}[filepath]    ${JSONData}[outputpath]    ${JSONData}[password]
    Log to file    Successfully converted '${JSONData}[filepath]' to '${JSONData}[outputpath]'


Convert PDF file to DOCX
    [Arguments]    ${TestData}    ${Element}
    ${Data}    	Split String	${TestData}	    ,
    ${JSONData}    Utilities.Convert String To Json    ${Data}
    Convert PDF to DOCX    ${JSONData}[filepath]    ${JSONData}[outputpath]    ${JSONData}[password]
    Log to file    Successfully converted '${JSONData}[filepath]' to '${JSONData}[outputpath]'


Convert DOCX file to PDF
    [Arguments]    ${TestData}    ${Element}
    ${Data}    	Split String	${TestData}	    ,
    ${JSONData}    Utilities.Convert String To Json    ${Data}
    Convert DOCX to PDF    ${JSONData}[filepath]    ${JSONData}[outputpath]    ${JSONData}[password]
    Log to file    Successfully converted '${JSONData}[filepath]' to '${JSONData}[outputpath]'


Move a file
    [Arguments]    ${TestData}    ${Element}
    ${Data}    	Split String	${TestData}	    ,
    ${JSONData}    Utilities.Convert String To Json    ${Data}
    Move File    ${JSONData}[sourcefilepath]    ${JSONData}[destinationpath]
    Log to file    Successfully moved '${JSONData}[sourcefilepath]' to '${JSONData}[destinationpath]'


Move a directory
    [Arguments]    ${TestData}    ${Element}
    ${Data}    	Split String	${TestData}	    ,
    ${JSONData}    Utilities.Convert String To Json    ${Data}
    Move Directory    ${JSONData}[sourcedirpath]    ${JSONData}[destinationpath]
    Log to file    Successfully moved '${JSONData}[sourcedirpath]' to '${JSONData}[destinationpath]'


Copy a file
    [Arguments]    ${TestData}    ${Element}
    ${Data}    	Split String	${TestData}	    ,
    ${JSONData}    Utilities.Convert String To Json    ${Data}
    Copy File    ${JSONData}[sourcefilepath]    ${JSONData}[destinationpath]
    Log to file    Successfully copied '${JSONData}[sourcefilepath]' to '${JSONData}[destinationpath]'


Copy a directory
    [Arguments]    ${TestData}    ${Element}
    ${Data}    	Split String	${TestData}	    ,
    ${JSONData}    Utilities.Convert String To Json    ${Data}
    Copy Directory    ${JSONData}[sourcedirpath]    ${JSONData}[destinationpath]
    Log to file    Successfully copied '${JSONData}[sourcedirpath]' to '${JSONData}[destinationpath]'


Capture screenshot
    [Arguments]    ${TestData}    ${Element}
    Capture Page Screenshot    ${TestData}
    Log to file    Captured screenshot and stored at '${TestData}'


Calculate file or directory size
    [Arguments]    ${TestData}    ${Element}
    ${Size}    Get size    ${TestData}
    Add To Json    file_or_dir_size    ${TestData}    ${Size}
    Log to file    File or directory size of '${TestData}' is '${Size}'


Get System Information
    ${SystemInfo}    Get System Info
    Add To Json    system_info    noKey    ${SystemInfo}
    Log to file    Retrieved system info '{SystemInfo}'


Write to a file
    [Arguments]    ${TestData}    ${Element}
    ${Data}    	Split String	${TestData}	    ,
    ${JSONData}    Utilities.Convert String To Json    ${Data}
    Create File    ${JSONData}[filepath]    ${JSONData}[content]
    Log to file    Content has been written to '${JSONData}[filepath]' successfully


Append to a file
    [Arguments]    ${TestData}    ${Element}
    ${Data}    	Split String	${TestData}	    ,
    ${JSONData}    Utilities.Convert String To Json    ${Data}
    Append to File    ${JSONData}[filepath]    ${JSONData}[content]
    Log to file    Content has been appended to '${JSONData}[filepath]' successfully


Read a file
    [Arguments]    ${TestData}    ${Element}
    Get File    ${TestData}
    Log to file    Content read from '${TestData}' successfully


Perform math operation
    [Arguments]    ${TestData}    ${Element}
    ${Result}    Evaluate Expression    ${TestData}
    Add To Json    math_results    result    ${Result}
    Log to file    Math operation result '${TestData} = ${Result}'


Terminate a process
    [Arguments]    ${TestData}    ${Element}
    Kill Process by Name    ${TestData}
    Log to file    Process '${TestData}' terminated


Check process status
    [Arguments]    ${TestData}    ${Element}
    ${Result}    Is Process Running    ${TestData}
    Add To Json    process_status    ${TestData}    ${Result}
    Log to file    Process '${TestData}' is '${Result}'


Get Process Information
    [Arguments]    ${TestData}    ${Element}
    ${Result}    Get Process Info    ${TestData}
    Add To Json    process_info    ${TestData}    ${Result}
    Log to file    Process Info '${TestData} : ${Result}'


Get all running processes
    ${Result}    List Running Processes
    Add To Json    all_running_processes    noKey    ${Result}
    Log to file    Retrieved all running processes


Execute python code
    [Arguments]    ${TestData}    ${Element}
    ${Result}    Evaluate    ${TestData}
    Add To Json    python_execution_results    ${TestData}    ${Result}
    Log to file    Python code executed '${TestData} : ${Result}'

Execute javascript code
    [Arguments]    ${TestData}    ${Element}
    ${Result}    Execute Javascript    ${TestData}
    Add To Json    javascript_execution_results    ${TestData}    ${Result}
    Log to file    JavaScript code executed '${TestData} : ${Result}'

Encrypt data
    [Arguments]    ${TestData}    ${Element}
    ${Result}    Encrypt Text with Keyfile    ${TestData}
    Add To Json    encrypted_data    ${TestData}    ${Result}
    Log to file    Encrypted text '${TestData} : ${Result}'


Decrypt data
    [Arguments]    ${TestData}    ${Element}
    ${Result}    Decrypt Text with Keyfile    ${TestData}
    Add To Json    decrypted_data    ${TestData}    ${Result}
    Log to file    Encrypted text '${TestData} : ${Result}'


Send Email with Attachments
    [Arguments]    ${TestData}    ${Element}
    ${Data}    	Split String	${TestData}	    ,
    ${JSONData}    Utilities.Convert String To Json    ${Data}
    Ibrog HTMLEmail AttachY    ${JSONData}['config_file']    ${JSONData}['reply_email']    ${JSONData}['to_email']    ${JSONData}['attachments']    ${JSONData}['subject']    ${JSONData}['body']
    Log to file    Email sent to '${Email}' with attachments


Send Email without Attachments
    [Arguments]    ${TestData}    ${Element}
    ${Data}    	Split String	${TestData}	    ,
    ${JSONData}    Utilities.Convert String To Json    ${Data}
    Ibrog HTMLEmail AttachN    ${JSONData}['config_file']    ${JSONData}['reply_email']    ${JSONData}['to_email']    ${JSONData}['subject']    ${JSONData}['body']
    Log to file    Email sent to '${Email}' without attachments


Compress Folder
    [Arguments]    ${TestData}    ${Element}
    ${Data}    	Split String	${TestData}	    ,
    ${JSONData}    Utilities.Convert String To Json    ${Data}
    Compress the Folder    ${JSONData}['folder_to_compress']    ${JSONData}['output_folder']
    Log to file     Folder '${JSONData}['folder_to_compress']' compressed and stored at location '${JSONData}['output_folder']'


Merge Excel files
    [Arguments]    ${TestData}    ${Element}
    ${Data}    	Split String	${TestData}	    ,
    ${JSONData}    Utilities.Convert String To Json    ${Data}
    DF to Excel     ${JSONData}['input_directory']    ${JSONData}['output_file']
    Log to file     Excel files in folder '${JSONData}['input_directory']' merged to '${JSONData}['output_file']'


Generate Fake Test Data
    [Arguments]    ${TestData}    ${Element}
    ${Data}    	Split String	${TestData}	    ,
    ${JSONData}    Utilities.Convert String To Json    ${Data}
    ${Result}    Generate Fake Data     ${JSONData}['data_type']    ${JSONData}['count']
    Add to Json    fake_test_data    data    ${Result}
    Log to file    '${JSONData}['count']' Fake data generated for '${JSONData}['data_type']'


# Condition statement to check the element is visible or not
IF_ElementVisible
    [Arguments]    ${Xpath}     ${Element}
    ${Visible}        Run Keyword and Return Status       element should be Visible         ${Xpath}
    [Return]        ${Visible}


# Looping statement based on elements visibilty
FOR_ElementVisible
    [Arguments]    ${Xpath}     ${Element}
    Scroll Element into View    ${Xpath}
    ${ElememtCount}        Get Element count       ${Xpath}

    [Return]        ${ElememtCount}

###****************************************************SERVER KEYWORDS***************************************************************####
Open Server Connection
    [Arguments]    ${IP_address}     ${port}        ${Element}
    Open Connection		${IP_address}		${port}
    Log to file    Opened Connection in Ip - ${Element}

Login To Server
    [Arguments]    ${user_id}     ${password}       ${Element}
	Login				${user_id}		${password}
    Log to file    Login to Server - ${Element}

Write To Server
    [Arguments]    ${command}       ${Element}
    Write		${command}
    Log to file    Write To Server - ${command}


Read Server Log
    [Arguments]    ${Element}
    ${logtxt}     Read
    Log to file    Read Server Log
    [Return]      ${logtxt}

Read Server Log Until
    [Arguments]    ${untilttxt}    ${Element}
    ${logtxt}     Read Until      ${untilttxt}
    Log to file    Read Server Log Until - ${untilttxt}
    [Return]      ${logtxt}


# To store the value of element (B)
Store element value
    [Arguments]    ${Xpath}    ${TestData}    ${Element}
    IF  "${AutomationType}" != "Mobile"
        Wait Until Element Is Visible    ${Xpath}    20s
        ${ElementValue}    Get Text    ${Xpath}
        ${Data}    	Split String	${TestData}	    ,
        ${JSONData}    Utilities.Convert String To Json    ${Data}
        Add To Json    variables    ${JSONData}[variablename]    ${ElementValue}    ${JSONData}[overwrite]
    END


# To create an excel file
Create Excel File
    [Arguments]    ${TestData}    ${Element}
    ${Data}    	Split String	${TestData}	    ,
    ${JSONData}    Utilities.Convert String To Json    ${Data}
    Create WorkBook    ${JSONData}[excelfilepath]    overwrite_file_if_exists=${JSONData}[overwritefileifexists]
    Close WorkBook
    Log to file    Created an Excel file '${JSONData}[excelfilepath]'


# To write data into excel cell
Write to Excel File
    [Arguments]    ${TestData}    ${Element}
    ${Data}    	Split String	${TestData}	    ,
    ${JSONData}    Utilities.Convert String To Json    ${Data}
    Open WorkBook    ${JSONData}[excelfilepath]
    Write to cell    ${JSONData}[cellname]    ${JSONData}[cellvalue]
    Save
    Close WorkBook
    Log to file    Inserted '${JSONData}[cellvalue]' into the cell named '${JSONData}[cellname]'

# To read excel data
Read from Excel File
    [Arguments]    ${TestData}    ${Element}
    ${Data}    	Split String	${TestData}	    ,
    ${JSONData}    Utilities.Convert String To Json    ${Data}
    Open WorkBook    ${JSONData}[excelfilepath]
    Switch Sheet    ${JSONData}[sheetname]
    ${ColumnNames}    Set Variable If    "${JSONData}[columnnames]".upper().strip() == "ALL"    None    ${JSONData}[columnnames]
    ${CellRange}    Set Variable If    "${JSONData}[cellrange]".upper().strip() == "ALL"    None    ${JSONData}[cellrange]
    ${SheetData}    Read sheet data    column_names=${ColumnNames}   cell_range=${CellRange}    get_column_names_from_header_row=True    trim=False
    Add To Json    excel_data   excel    ${SheetData}
    Close WorkBook
    Log to file    Data retrieved from the Excel file located at '${JSONData}[excelfilepath]'


###****************************************************MAIN KEYWORDS***************************************************************####

# To handle POM (F)
POM Handler
    IF     "${DataSource}".upper().strip() == "EXCEL"
        Switch Sheet        POM
        ${SheetData}        Read Sheet Data             get_column_names_from_header_row=True       trim=True
    ELSE
        ${SheetData}    Get Test Data From DB    POM
    END

    ${Locators}        Create dictionary

    FOR     ${row}       IN     @{sheetdata}
        IF      "${AutomationType}"=="Mobile"
        #     Set to dictionary        ${Locators}        ${row['locator Name']}=${row['Type']}=${row['Value']}        ## Username=id:impinituser
        # ELSE IF     "${row['Type']}"=="xpath"
            Set to dictionary        ${Locators}        ${row['locator Name']}=${row['Value']}        ## Username=id:impinituser
        ELSE
            Set to dictionary        ${Locators}        ${row['locator Name']}=${row['Type']}:${row['Value']}        ## Username=id:impinituser
        END
    END
    Set Global Variable        ${Locators}


# To segregate executable and non-executable test cases (F)
Execution Handler
    [Arguments]    ${Group}
    IF     "${DataSource}".upper().strip() == "EXCEL"
        Switch Sheet        Execution Handler
        ${SheetData}        Read Sheet Data        get_column_names_from_header_row=True
    ELSE
        ${SheetData}    Get Test Data From DB    TestRepo
    END

    ${ExecuteTestCaseList}        Create List
    ${ExecuteTestCaseData}        Create List
    Set Global Variable    ${Group}

    FOR        ${row}        IN     @{sheetdata}
        IF    "${row['Group']}"=="None"
            Set to Dictionary    ${row}    Group    G1
        END
        IF     "${row['Execution Flag']}".upper().strip() =="YES" and "${row['Group']}".upper().strip() =="${Group}"
            Append to List       ${ExecuteTestCaseData}       ${row}
            Append to List       ${ExecuteTestCaseList}       ${row['Test Case ID']}
        END
    END
    IF     len(${ExecuteTestCaseData}) !=0
        Set Global Variable        ${ExecuteTestCaseData}
        Set Global Variable        ${ExecuteTestCaseList}
        Initial Test Setup    ${Group}
        Set Global Variable    ${SkipTeardown}    False
    ELSE
        Set Global Variable    ${SkipTeardown}    True
        Pass Execution    Group ${Group} not present
    END


# To handle test cases (F)
Test Case Handler
    IF    "${DataSource}".upper().strip() == "EXCEL"
        Switch Sheet        Test Case Handler
        ${SheetData}        Read Sheet Data        get_column_names_from_header_row=True
        Close Workbook
    ELSE
        ${SheetData}    Get Test Data From DB    TestSteps
    END
    Create File    ${CURDIR}\\WorkBook Check Points\\${TEST NAME}_WorkBook_Closed.txt

    ${AllTestCaseList}        Create Dictionary
    FOR        ${row}        IN     @{sheetdata}
        IF     "${row['Test Case ID']}"!="None"
            ${TempDict}        Create dictionary        TestCaseID=${row['Test Case ID']}        Description=${row['Description']}    Keywords=@{EMPTY}
            Set To Dictionary        ${AllTestCaseList}        ${row['Test Case ID']}=${TempDict}
        END
        ${Buffer}        Set Variable        ${EMPTY}
        ${List}        Create List        ${row['Pre Condition']}        ${row['Steps']}        ${row['WebElement']}        ${row['Test Data']}        ${row['Sleep']}        ${row['Execution Flag']}    ${row['Expected Result']}        ${Buffer}
        Append to List        ${AllTestCaseList['${TempDict['TestCaseID']}']['Keywords']}        ${List}
    END
    Set Global Variable        ${AllTestCaseList}


# To handle test suite info (F)
Test Suite Info Handler
    Switch Sheet        Test Suite Info
    ${SheetData}        Read Sheet Data        get_column_names_from_header_row=True
    ${SuiteinfoList}        Evaluate        $sheetdata[0]
    Set Global Variable        ${SuiteinfoList}

    # Generic info
    Set Global Variable    ${AutomationType}    ${SuiteinfoList['Automation Type']}
    Set Global Variable    ${ApplicationName}    ${SuiteinfoList['Application Name']}
    Set Global Variable    ${ApplicationVersion}    ${SuiteinfoList['Application Version']}
    Set Global Variable    ${OperatingSystem}    ${SuiteinfoList['Operating System']}
    Set Global Variable    ${Environment}    ${SuiteinfoList['Environment']}
    Set Global Variable    ${SendEmail}    ${SuiteinfoList['Receive Report via Email?']}
    Set Global Variable    ${Email}    ${SuiteinfoList['Email']}
    Set Global Variable    ${RegressionTest}    ${SuiteinfoList['Is this Regression test?']}
    Set Global Variable    ${GenerateReport}    ${SuiteinfoList['Generate daily report']}
    Set Global Variable    ${ReportPath}    ${SuiteinfoList['Report Path']}
    Set Global Variable    ${RecordScreen}    ${SuiteinfoList['Video Recording']}
    Set Global Variable    ${DataSource}    ${SuiteinfoList['Data Source']}
    Set Global Variable    ${VideoRecPath}    ${SuiteinfoList['Video Recording path']}

    IF    "${ReportPath}".strip() == "" or "${ReportPath}".strip() == "None"
        Fail    The report path must not be left blank. Kindly specify a directory path in the Excel sheet for storing the reports.
    END

    IF     "${AutomationType}" == "Web"
        Set Global Variable    ${URL}    ${SuiteinfoList['URL']}
        Set Global Variable    ${Browser}    ${SuiteinfoList['Browser']}
    ELSE
        Set Global Variable    ${OperatingSystemVersion}    ${SuiteinfoList['Operating System Version']}
    END


# To execute test cases (F)
Execute Test Cases
    IF      "${AutomationType}"=="Mobile"
        Import Library      AppiumLibrary
    ELSE
        Import Library      Selenium2Library
    END

    FOR    ${index}    ${TestCaseID}    IN ENUMERATE      @{ExecuteTestCaseList}
        Log to file     Testcase '${TestCaseID}' execution started
        Set Global Variable    ${SkipTestCase}    False

        Set Global Variable     ${ExecutionBool}        None
        Set Global Variable     ${FunctionBool}     False
        ${ForLoopList}    Create List
        Set Global Variable     ${ForLoopList}


        FOR        ${KeywordData}        IN        @{AllTestCaseList['${TestCaseID}']['Keywords']}
            IF     "${KeywordData[-3]}"=="None"
                No Operation
            ELSE IF      "${KeywordData[-3].upper().strip()}"=="YES"
                IF   "${SkipTestCase}"=="False"
                    Set Global Variable    ${CurrentTestCaseId}    ${TestCaseID}

                    # IF    "${RecordScreen}".upper().strip() == "YES"
                    #     ADB Start Screen Recording    //sdcard//screenrecord.mp4
                    #     Set Global Variable    ${ScreenRecordingStarted}    True
                    # END

                    IF      "${FunctionBool}"=="False" and "${ExecutionBool}"=="None"
                        ${ExecutionBool}     Execute Keyword        @{KeywordData}    ${TestCaseID}
                        Set Global Variable     ${ExecutionBool}
                    END

                    IF    "${SkipTestCase}" != "Skipped"
                        IF     "${KeywordData[1].upper()}" in ${FunctionKeywords}
                            Set Global Variable     ${FunctionBool}     True

                            IF      "${KeywordData[1]}"=="FOR_ElementVisible"
                                IF      "${KeywordData[3]}"!="None"
                                    ${forloopData}    	Split String	${KeywordData[3].upper()}	    ,
                                    ${JSONData}     Utilities.Convert String To Json    ${forloopData}

                                    ${StartRange}       Set Variable        ${JSONData}[startrange]
                                    ${StepValue}        Set Variable        ${JSONData}[stepvalue]
                                ELSE IF     "${KeywordData[3]}"=="None"
                                    ${StartRange}       Set Variable        1
                                    ${StepValue}        Set Variable        1
                                END
                            END
                        END

                        IF     "${FunctionBool}"=="True" and "${KeywordData[1].upper()}" not in ${FunctionKeywords}
                            IF      "${ExecutionBool}"=="True"
                                ${ExecutionBool_1}     Execute Keyword        @{KeywordData}    ${TestCaseID}

                            ELSE IF     "${ExecutionBool}" not in ['None', 'True', 'False']
                                Append To List    ${ForLoopList}    ${KeywordData}
                                # log     ${ForLoopList}      WARN

                            END

                        ELSE IF     "${FunctionBool}"=="True"
                            IF     "${KeywordData[1].upper()}"=="ELSE"
                                ${ExecutionBool}     Set Variable IF        "${ExecutionBool}"=="True"        False     True
                                Set Global Variable     ${ExecutionBool}
                            ELSE IF     "${KeywordData[1].upper()}"=="END"
                                IF      len($ForLoopList)!=0
                                    FOR     ${i}        IN RANGE    ${StartRange}        ${ExecutionBool}+1      ${StepValue}
                                        FOR     ${LoopData}     IN      @{ForLoopList}
                                            ${indexvalue_xpath}        Check i value in Xpath      ${Locators['${LoopData[2]}']}       ${i}
                                            # ${ExecutionBool_1}     Execute Keyword        @{LoopData}    ${TestCaseID}
                                            ${ExecutionBool_1}     Execute Keyword        ${LoopData[0]}        ${LoopData[1]}      ${LoopData[2]}     ${LoopData[3]}      ${LoopData[4]}        ${LoopData[5]}        ${LoopData[6]}      ${indexvalue_xpath}        ${TestCaseID}
                                                                        #[Arguments]        ${PreCondition}        ${Steps}        ${Element}               ${TestData}        ${Sleep}            ${ExecutionFlag}        ${ExpectedResult}    ${TestCaseID}
                                    	END
                                    END
                                END
                                Set Global Variable     ${ExecutionBool}        None
                                Set Global Variable     ${FunctionBool}     False
                                ${ForLoopList}    Create List
                                Set Global Variable     ${ForLoopList}

                                No Operation
                            END
                        END
                    END
                    Remove Temp Files    Temp
                END
            END
        END

        IF    "${SkipTestCase}"=="False"
            Log to file    Testcase '${TestCaseID}' execution ended
        ELSE IF     "${SkipTestCase}"=="True"
            Log to file    Testcase '${TestCaseID}' failed
        END

        IF    len(${CurrentTestCaseAssertionFails})!=0
            Set to Dictionary    ${TempReportData}    Status    FAILED
            ${NewReason}    Evaluate    ", ".join(${CurrentTestCaseAssertionFails})
            ${OldReason}    Set Variable    ${TempReportData}[Reason]

            IF    """${OldReason}""" != """None"""
                ${CombinedReason}    Catenate    SEPARATOR=,    ${NewReason}    ${OldReason}
                ${NewErrorType}    Set variable    Assertion/Script Error
            ELSE
                ${CombinedReason}    Set Variable    ${NewReason}
                ${NewErrorType}    Set variable    Assertion Error
            END

            Set to Dictionary    ${TempReportData}    Reason    ${CombinedReason}
            Set to Dictionary    ${TempReportData}    ErrorType    ${NewErrorType}
        END

        Append to list    ${ReportData}    ${TempReportData}

        FOR    ${errors}    IN    @{CurrentTestCaseAssertionFails}
            Append To List    ${TotalAssertionFails}    ${errors}
        END
        Run Keyword If    len(${CurrentTestCaseAssertionFails})!=0    Evaluate	$CurrentTestCaseAssertionFails.clear()
    END
    Create File    ${CURDIR}\\Execution Check Points\\${TEST NAME}_Execution_Completed.txt
    # Run Keyword If    ${ScreenRecordingStarted}    ADB Stop Screen Recording    //sdcard//screenrecord.mp4    ${ReportPath}\\Execution Summary\\Execution_Summary_${TimeStamp}\\Group_${Group}\\Videos\\screenrecord.mp4


Convert To Xpath
    [Arguments]        ${TestData}
    ${TestData}    	Split String	${TestData}	    ,
    ${JSONData}     Utilities.Convert String To Json    ${TestData}
    IF      'text' in ${JSONData}
        ${testdata_xpath}       Set Variable        ${JSONData}[text]
        ${botXpath}     Evaluate        '//*[.="'+$testdata_xpath+'"]'
    END
    [Return]        ${botXpath}

# To execute keywords (F)
Execute Keyword
    [Arguments]        ${PreCondition}        ${Steps}        ${Element}        ${TestData}        ${Sleep}        ${ExecutionFlag}        ${ExpectedResult}    ${Buffer}        ${TestCaseID}
    IF     "${ExecutionFlag}".upper().strip() == "YES"
        ${value_0}      Set Variable        None
        IF      "${FunctionBool}"=="True" and len($ForLoopList)!=0
            ${Element_xpath}      Set Variable        ${Buffer}
            # ${Element}      Set Variable        ${Element}
        ELSE IF     """${Element}"""!="""None"""
            ${Element_xpath}      Set Variable        ${Locators['${Element}']}
        END

        TRY
            ${KeywordStartTime}    Get Current Date    result_format=%H:%M:%S
            Set Global Variable    ${AssertionError}    False
            ${Status}    Set Variable    PASSED
            ${Reason}    Set Variable    ${EMPTY}
            ${ErrorType}    Set Variable    ${EMPTY}
            ${AddStepToDict}    Set Variable    True


            IF      "${Steps}"=="None"
                Set Global Variable    ${SkipTestCase}    Skipped
                Fail     Test case '${TestCaseID}' is skipped
            END

            IF     "${PreCondition}"!="None"
                FOR        ${KeywordData}        IN        @{AllTestCaseList['${PreCondition}']['Keywords']}
                    ${ExecutionFileExist}    Run keyword and Ignore Error    OperatingSystem.File Should Exist    ${CURDIR}\\Temp\\Precondition_Executing.txt
                    IF     "${ExecutionFileExist}[0]" == "FAIL"    Create File    ${CURDIR}\\Temp\\Precondition_Executing.txt
                    Execute Keyword        @{KeywordData}    ${PreCondition}
                END
            END

            ${FailedFileExist}    Run keyword and Ignore Error    OperatingSystem.File Should Not Exist    ${CURDIR}\\Temp\\Precondition_Failed.txt
            IF     "${FailedFileExist}[0]" == "PASS"
                IF        "${Steps.upper()}" in ${1_arglist}
                    IF      """${TestData}"""=="""None"""
                        ${argum}         Set Variable       ${Element_xpath}
                    ELSE
                        ${argum}         Set Variable       ${TestData}
                        IF      "${Steps.upper()}"=="CLICK"
                            ${argum}        Convert To Xpath         ${argum}
                        END
                    END
                    ${value_0}      Keyword_arg1        ${Steps}        ${argum}        ${Element}
                ELSE IF      "${Steps.upper()}" in ${2_arglist}
                    Keyword_arg2        ${Steps}        ${Element_xpath}        ${TestData}        ${Element}
                ELSE IF      "${Steps.upper()}" in ${0_arglist}
                    ${value_0}      Keyword_arg0        ${Steps}
                ELSE IF      "${Steps.upper()}" in ${3_arglist}
                    Keyword_arg3        ${Steps}        ${Element_xpath}       ${TestData}        ${Element}
                ELSE IF      "${Steps.upper()}" in ${Verifylist}
                    Keyword_arg2        ${Steps}        ${Element_xpath}        ${ExpectedResult}        ${Element}
                ELSE IF    "${Steps.upper()}" in ${Exceptionlist1}
                    Keyword_arg1        ${Steps}        ${ExpectedResult}        ${Element}
                ELSE IF    "${Steps.upper()}" in ${Exceptionlist2}
                    Keyword_arg4        ${Steps}        ${Element_xpath}        ${TestData}        ${ExpectedResult}        ${Element}
                ELSE IF    "${Steps.upper()}" in ${FunctionKeywords}
                    ${value_0}      FunctionKeyword_arg1        ${Steps}        ${Element_xpath}        ${Element}
                ELSE
                    Fail   Keyword '${Steps}' not found
                END
                # Induce Sleep
                Run Keyword if         "${Sleep}"!="None"        Sleep        ${Sleep}
            ELSE
                ${AddStepToDict}    Set Variable    False
            END

        EXCEPT    AS    ${Exception}
            ${Time}    Get Current Date    result_format=%d_%m_%Y_%H_%M_%S
            Log to file     ${Exception}     ERROR
            Capture Page Screenshot    ${ReportPath}\\Execution Summary\\Execution_Summary_${TimeStamp}\\Group_${Group}\\Snapshots\\Error_Snap_${Time}.png
            Log to file    Error screenshot captured
            ${Status}    Set Variable      FAILED
            ${Reason}    Set Variable    ${Exception}

            IF    "${AssertionError}"!="True" and "${SkipTestCase}" != "Skipped"
                Set Global Variable    ${SkipTestCase}    True
                Append To List    ${TotalScriptFails}    ${Exception}
                ${ExecutionFileExist}    Run keyword and Ignore Error    OperatingSystem.File Should Exist    ${CURDIR}\\Temp\\Precondition_Executing.txt
                IF     "${ExecutionFileExist}[0]" == "PASS"
                    Create File    ${CURDIR}\\Temp\\Precondition_Failed.txt
                END
            END
        END
        ${KeywordEndTime}    Get Current Date    result_format=%H:%M:%S
        ${KeywordDuration}    Subtract Time From Time		${KeywordEndTime}		${KeywordStartTime}
        ${KeywordDuration}    Convert To Integer    ${KeywordDuration}
        ${KeywordStartTime}    Convert To 12hr Format    ${KeywordStartTime}


        IF    "${Status}"!="PASSED"
            IF    ${AssertionError}
                ${ErrorType}    Set Variable    Assertion Error
            ELSE IF    "${SkipTestCase}" == "Skipped"
                ${ErrorType}    Set Variable    Syntax Error
                ${Status}    Set Variable    SKIPPED
                ${Reason}    Set Variable     Data Missing in Excel sheet
            ELSE
                ${ErrorType}    Set Variable    Script Error
            END
        END

        IF    ${AddStepToDict}
            IF    "${Element}" != "None"
                ${Steps}   Set variable    ${Steps} - ${Element}
            END
            ${TempStepWiseReportData}    Create Dictionary    SerialNum=${EMPTY}    TestCaseID=${CurrentTestCaseId}    TestCaseDesc=${AllTestCaseList}[${CurrentTestCaseId}][Description]    StepNum=${EMPTY}    Steps=${Steps}    ExecutionStartTime=${KeywordStartTime}    Duration=${KeywordDuration} sec        Status=${Status}    ErrorType=${ErrorType}    Reason=${Reason}
            Append to list     ${StepWiseReportData}    ${TempStepWiseReportData}
            IF    ${AssertionError}
                ${Reason}    Set Variable     None
            END
            ${TempReportData}    Create Dictionary    SerialNum=${EMPTY}    TestCaseID=${CurrentTestCaseId}    TestCaseDesc=${AllTestCaseList}[${CurrentTestCaseId}][Description]    ExecutionStartTime=${EMPTY}    Duration=${EMPTY}    Status=${Status}    ErrorType=${ErrorType}    Reason=${Reason}
            Set Global Variable    ${TempReportData}
        END
        RETURN        ${value_0}
    END

Check i value in Xpath
    [Arguments]     ${xpath}        ${loopindex}
    ${loopindex}        Convert To String       ${loopindex}
    ${if_i_present}     Evaluate      re.search("[ind]",$xpath)
    IF      ${if_i_present}
        ${new_xpath}        Evaluate      re.sub("ind",$loopindex,$xpath)      #${xpath}        $i        ${loopindex}
    ELSE IF     ${loopindex}"=="None"
        ${new_xpath}        Set Variable        ${xpath}
    END
    [Return]        ${new_xpath}


# To wait until the file is created (F)
Wait until the file is created
    [Arguments]    ${FileName}
    Run keyword and Ignore Error    Wait Until Created    ${CURDIR}\\WorkBook Check Points\\${FileName}_WorkBook_Closed.txt    15s


# To check if the execution is completed or not (F)
Check if execution completed
    ${Files}    OperatingSystem.List files in directory    ${CURDIR}\\Execution Check Points
    ${ExecutionCompletionCount}    Get Length    ${Files}
    ${SuiteData}    Load Json From File    ${CURDIR}\\Test Suite Info\\TestSuiteInfo.json
    IF    ${ExecutionCompletionCount} == ${SuiteData}[browser_count]
        RETURN    True    ${SuiteData}[browser_count]
    ELSE
        RETURN    False    ${SuiteData}[browser_count]
    END

*** Test Cases ***
# Suite setup
Setup
    ${TestSuiteStartTime}    Get Current Date     result_format=%H:%M:%S
    ${TimeStamp}    Get Current Date    result_format=%d_%m_%Y_%H_%M_%S
    Create File    ${CURDIR}\\Test Suite Info\\TestSuiteInfo.json    {"report_dir_time_stamp": "${TimeStamp}", "suite_start_time" : "${TestSuiteStartTime}","browser_count": 0, "group_g1": {"total_testcases_count": 0,"failed_testcase_count": 0,"passed_testcase_count": 0,"skipped_testcase_count": 0,"total_scriptfail_count": 0,"total_assertionfail_count": 0,"total_fail_count": 0,"duration": 0, "failed_testcase_info": null},"group_g2": {"total_testcases_count": 0,"failed_testcase_count": 0,"passed_testcase_count": 0,"skipped_testcase_count": 0,"total_scriptfail_count": 0,"total_assertionfail_count": 0,"total_fail_count": 0,"duration": 0, "failed_testcase_info": null},"group_g3": {"total_testcases_count": 0,"failed_testcase_count": 0,"passed_testcase_count": 0,"skipped_testcase_count": 0,"total_scriptfail_count": 0,"total_assertionfail_count": 0,"total_fail_count": 0,"duration": 0, "failed_testcase_info": null},"group_g4": {"total_testcases_count": 0,"failed_testcase_count": 0,"passed_testcase_count": 0,"skipped_testcase_count": 0,"total_scriptfail_count": 0,"total_assertionfail_count": 0,"total_fail_count": 0,"duration": 0, "failed_testcase_info": null},"group_g5": {"total_testcases_count": 0,"failed_testcase_count": 0,"passed_testcase_count": 0,"skipped_testcase_count": 0,"total_scriptfail_count": 0,"total_assertionfail_count": 0,"total_fail_count": 0,"duration": 0, "failed_testcase_info": null}, "variables": {}, "excel_data": {}, "query_results": {}, "file_or_dir_size": {}, "system_info": null, "math_results": {}, "process_status": {}, "process_info": {}, "all_running_processes": null, "python_execution_results": {}, "javascript_execution_results": {}, "encrypted_data": {}, "decrypted_data": {}, "fake_test_data": {}}
    Create Directory    ${CURDIR}\\VideoRec_Info
    Remove Temp Files    Temp,WorkBook Check Points,Execution Check Points,VideoRec_Info

    Open Workbook        ${IBorgTemplatePath}
    Switch Sheet        Execution Handler
    ${SheetData}        Read Sheet Data        get_column_names_from_header_row=True
    Create File    ${CURDIR}\\WorkBook Check Points\\${TEST NAME}_WorkBook_Closed.txt
    ${GroupCount}    Evaluate    set(row['Group'] for row in ${SheetData} if row['Execution Flag'].upper().strip() == "YES")
    Evaluate    $GroupCount.discard(None)
    ${BrowserCount}    Evaluate    len(${GroupCount})
    Add To Json    browser_count    nokey    ${BrowserCount}

    Test Suite Info Handler
    Close workbook

    ${RecordPath}       Set Variable        ${ReportPath}\\Execution Summary\\Execution_Summary_${TimeStamp}\\screenrecord.mp4
    ${RecordPath}        Normalize Path        ${RecordPath}
    Set Global Variable        ${RecordPath}
    IF      "${RecordScreen}"=="Yes"
        Create File     ${CURDIR}\\VideoRec_Info\\StartRecording.txt        "${RecordPath}"
        Import Library      D:\\New Framework\\Newfwrk_V2.0\\videorec.py
    END

# For group G1 execution
IBorgDataDrivenControl1
    Wait until the file is created    Setup
    Open Workbook        ${IBorgTemplatePath}
    Test Suite Info Handler
    POM Handler
    Execution Handler    G1
    Test Case Handler
    Execute Test Cases
    [Teardown]    Teardown

# For group G2 execution
IBorgDataDrivenControl2
    Wait until the file is created    IBorgDataDrivenControl1
    Open Workbook        ${IBorgTemplatePath}
    Test Suite Info Handler
    POM Handler
    Execution Handler    G2
    Test Case Handler
    Execute Test Cases
    [Teardown]    Teardown

# For group G3 execution
IBorgDataDrivenControl3
    Wait until the file is created    IBorgDataDrivenControl2
    Open Workbook        ${IBorgTemplatePath}
    Test Suite Info Handler
    POM Handler
    Execution Handler    G3
    Test Case Handler
    Execute Test Cases
    [Teardown]    Teardown


# For group G4 execution
IBorgDataDrivenControl4
    Wait until the file is created    IBorgDataDrivenControl3
    Open Workbook        ${IBorgTemplatePath}
    Test Suite Info Handler
    POM Handler
    Execution Handler    G4
    Test Case Handler
    Execute Test Cases
    [Teardown]    Teardown

# For group G5 execution
IBorgDataDrivenControl5
    Wait until the file is created    IBorgDataDrivenControl4
    Open Workbook        ${IBorgTemplatePath}
    Test Suite Info Handler
    POM Handler
    Execution Handler    G5
    Test Case Handler
    Execute Test Cases
    [Teardown]    Teardown
