Attribute VB_Name = "Libraries"
Sub LastUsage()
    ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range("B7").Value = DateTime.Now
    ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range("B7").NumberFormat = "hh:mm:ss AM/PM"
End Sub

Sub Initialize()
    track_changes = False
    Application.EnableEvents = True
    custom_field_2_count = 0
    custom_field_1_count = 0
    fix_version_count = 0
    issue_links_count = 0
    filter_search_query = ""
    
    PopulateOpenStates
    
    Set scriptControl = CreateObject("MSScriptControl.ScriptControl")
    scriptControl.Language = "JScript"
    
    Set table_range = ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range(QUERY_UPDATE_HEADER_RANGE)
    Set holidays = ThisWorkbook.Worksheets(SHEET_HOLIDAYS).ListObjects("Table_IndiaHoliday2019")
    ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range(QUERY_UPDATE_HEADER_RANGE).Interior.ColorIndex = 0
    If ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Evaluate("ISREF(JiraQueryUpdateTable)") Then
        Set query_update_table = ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).ListObjects("JiraQueryUpdateTable")
        'Avoid run-time error if attempted on empty table
        If Not query_update_table.DataBodyRange Is Nothing Then
            query_update_table.DataBodyRange.Interior.ColorIndex = xlNone
            query_update_table.AutoFilter.ShowAllData
            query_update_table.DataBodyRange.Delete
        End If
    Else
        WriteHeader
    End If
    
    With ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).ListObjects("JiraQueryUpdateTable")
        'Constants are zero indexed, hence +1 for each column number
        .ListColumns(ISSUE_KEY_COLUMN + 1).Range.ColumnWidth = 10 'ID
        .ListColumns(ISSUE_TYPE_COLUMN + 1).Range.ColumnWidth = 10 'Issue type
        .ListColumns(ISSUE_STATUS_COLUMN + 1).Range.ColumnWidth = 10 'Status
        .ListColumns(ISSUE_SUMMARY_COLUMN + 1).Range.ColumnWidth = 60 'Summary
        .ListColumns(ISSUE_ASSIGNEE_COLUMN + 1).Range.ColumnWidth = 10 'Assignee
        .ListColumns(ISSUE_EPIC_LINK_COLUMN + 1).Range.ColumnWidth = 10 'Epic key
        .ListColumns(ISSUE_BLOCKED_BY_COLUMN + 1).Range.ColumnWidth = 10 'Blocked by
        .ListColumns(ISSUE_BLOCKS_COLUMN + 1).Range.ColumnWidth = 10 'Blocks
        .ListColumns(ISSUE_FIX_VERSION_COLUMN + 1).Range.ColumnWidth = 20 'Fix version(s)
        .ListColumns(ISSUE_PRIORITY_COLUMN + 1).Range.ColumnWidth = 10 'Priority
        .ListColumns(ISSUE_CUSTOM_FIELD_2_COLUMN + 1).Range.ColumnWidth = 10 'Custom Field 2
        .ListColumns(ISSUE_DUE_DATE_COLUMN + 1).Range.ColumnWidth = 10 'Due date
        .ListColumns(ISSUE_START_DATE_COLUMN + 1).Range.ColumnWidth = 10 'Start date
        .ListColumns(ISSUE_END_DATE_COLUMN + 1).Range.ColumnWidth = 10 'End date
        .ListColumns(ISSUE_CUSTOM_FIELD_1_COLUMN + 1).Range.ColumnWidth = 10 'Custom Field 1
        .ListColumns(ISSUE_COMPONENTS_COLUMN + 1).Range.ColumnWidth = 10 'Component(s)
        .ListColumns(ISSUE_ORIGINAL_ESTIMATE_COLUMN + 1).Range.ColumnWidth = 10 'Original estimate
        .ListColumns(ISSUE_REMAINING_ESTIMATE_COLUMN + 1).Range.ColumnWidth = 10 'Remaining estimate
        .ListColumns(ISSUE_TIME_SPENT_COLUMN + 1).Range.ColumnWidth = 10 'Time spent/worklog
        .ListColumns(ISSUE_COMMENTS_COLUMN + 1).Range.ColumnWidth = 60 'Comments
        .ListColumns(ISSUE_BANDWIDTH_COLUMN + 1).Range.ColumnWidth = 10 'Bandwidth
        .ListColumns(ISSUE_TREND_COLUMN + 1).Range.ColumnWidth = 17 'Trend
        .ListColumns(ISSUE_ASSESSMENT_COLUMN + 1).Range.ColumnWidth = 20 'Assessment
        .ListColumns(ISSUE_BEST_END_COLUMN + 1).Range.ColumnWidth = 10 'Best end date
        .ListColumns(ISSUE_DUE_DATE_COLUMN + 1).Range.NumberFormat = "dd-mmm-yy" 'Due date
        .ListColumns(ISSUE_START_DATE_COLUMN + 1).Range.NumberFormat = "dd-mmm-yy" 'Start date
        .ListColumns(ISSUE_END_DATE_COLUMN + 1).Range.NumberFormat = "dd-mmm-yy" 'End date
        .ListColumns(ISSUE_BEST_END_COLUMN + 1).Range.NumberFormat = "dd-mmm-yy" 'Best end date
        .ListColumns(ISSUE_BANDWIDTH_COLUMN + 1).Range.NumberFormat = "0.00%" 'Bandwidth
    End With
End Sub

Sub PopulateOpenStates()
    Set dict_open_states = Nothing
    dict_open_states("New") = True
    dict_open_states("Open") = True
    dict_open_states("Ready") = True
    dict_open_states("Reopened") = True
    dict_open_states("In Analysis") = True
    dict_open_states("In Review") = True
    dict_open_states("In Testing") = True
    dict_open_states("In Progress") = True
    dict_open_states("In Design") = True
    dict_open_states("In Development") = True
End Sub

Sub PopulateColumnUpdateType()
    update_type(ISSUE_KEY_COLUMN) = PUT_POST_NONE 'ID
    update_type(ISSUE_TYPE_COLUMN) = PUT_POST_NONE 'Issue type
    update_type(ISSUE_STATUS_COLUMN) = POST_STATUS 'Status
    update_type(ISSUE_SUMMARY_COLUMN) = PUT_GENERIC_FIELD 'Summary
    update_type(ISSUE_ASSIGNEE_COLUMN) = PUT_GENERIC_FIELD 'Assignee
    update_type(ISSUE_EPIC_LINK_COLUMN) = PUT_GENERIC_FIELD 'Epic link
    update_type(ISSUE_BLOCKED_BY_COLUMN) = PUT_GENERIC_UPDATE 'Blocked by
    update_type(ISSUE_BLOCKS_COLUMN) = PUT_GENERIC_UPDATE 'Blocks
    update_type(ISSUE_FIX_VERSION_COLUMN) = PUT_GENERIC_FIELD 'Fix version(s)
    update_type(ISSUE_PRIORITY_COLUMN) = PUT_GENERIC_UPDATE 'Priority
    update_type(ISSUE_CUSTOM_FIELD_2_COLUMN) = PUT_GENERIC_FIELD 'Custom Field 2
    update_type(ISSUE_DUE_DATE_COLUMN) = PUT_GENERIC_FIELD 'Due date
    update_type(ISSUE_START_DATE_COLUMN) = PUT_GENERIC_FIELD 'Start date
    update_type(ISSUE_END_DATE_COLUMN) = PUT_GENERIC_FIELD 'End date
    update_type(ISSUE_CUSTOM_FIELD_1_COLUMN) = PUT_GENERIC_FIELD 'Custom Field 1
    update_type(ISSUE_COMPONENTS_COLUMN) = PUT_GENERIC_FIELD 'Component(s)
    update_type(ISSUE_ORIGINAL_ESTIMATE_COLUMN) = PUT_GENERIC_UPDATE 'Original estimate
    update_type(ISSUE_REMAINING_ESTIMATE_COLUMN) = PUT_GENERIC_UPDATE_REM_ESTIMATE 'Remaining Estimate
    update_type(ISSUE_TIME_SPENT_COLUMN) = POST_WORKLOG 'Time spent/worklog
    update_type(ISSUE_COMMENTS_COLUMN) = PUT_GENERIC_UPDATE 'Comment
End Sub

'Avoid holding resources after exit
Sub Cleanup()
    SESSION_ID = ""
    Set jira_json = Nothing
    Set jira_issue_json = Nothing
    Set jira_json_fields = Nothing
    Set JiraService = Nothing
    Set scriptControl = Nothing
    LastUsage
    ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range("B6").Value = DateDiff("s", start_time, DateTime.Now) & " sec"
    track_changes = True
End Sub

'Header for query and update table
Sub WriteHeader()
    ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range(QUERY_UPDATE_DATABODYRANGE_COL & ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).UsedRange.Rows.Count).Delete
    Dim HEADER(0, QUERY_UPDATE_COLUMN_COUNT) As Variant
    HEADER(0, ISSUE_KEY_COLUMN) = "ID"
    HEADER(0, ISSUE_TYPE_COLUMN) = "Issue Type"
    HEADER(0, ISSUE_STATUS_COLUMN) = "Status"
    HEADER(0, ISSUE_SUMMARY_COLUMN) = "Summary"
    HEADER(0, ISSUE_ASSIGNEE_COLUMN) = "Assignee"
    HEADER(0, ISSUE_EPIC_LINK_COLUMN) = "Epic Link"
    HEADER(0, ISSUE_BLOCKED_BY_COLUMN) = "Blocked By"
    HEADER(0, ISSUE_BLOCKS_COLUMN) = "Blocks"
    HEADER(0, ISSUE_FIX_VERSION_COLUMN) = "Fix version(s)"
    HEADER(0, ISSUE_PRIORITY_COLUMN) = "Priority"
    HEADER(0, ISSUE_CUSTOM_FIELD_2_COLUMN) = "Custom Field 2"
    HEADER(0, ISSUE_DUE_DATE_COLUMN) = "Due Date"
    HEADER(0, ISSUE_START_DATE_COLUMN) = "Start Date"
    HEADER(0, ISSUE_END_DATE_COLUMN) = "End Date"
    HEADER(0, ISSUE_CUSTOM_FIELD_1_COLUMN) = "Custom Field 1"
    HEADER(0, ISSUE_COMPONENTS_COLUMN) = "Component(s)"
    HEADER(0, ISSUE_ORIGINAL_ESTIMATE_COLUMN) = "Original Estimate"
    HEADER(0, ISSUE_REMAINING_ESTIMATE_COLUMN) = "Remaining Estimate"
    HEADER(0, ISSUE_TIME_SPENT_COLUMN) = "Time Spent"
    HEADER(0, ISSUE_COMMENTS_COLUMN) = "Add Comment"
    HEADER(0, ISSUE_BANDWIDTH_COLUMN) = "Bandwidth Planned"
    HEADER(0, ISSUE_TREND_COLUMN) = "Trend"
    HEADER(0, ISSUE_BEST_END_COLUMN) = "Best Case End Date"
    HEADER(0, ISSUE_ASSESSMENT_COLUMN) = "Assessment"
    table_range.Rows(1).Value = HEADER()
    'Format as a table
    Dim query_update_table As ListObject
    Set query_update_table = ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).ListObjects.Add(xlSrcRange, table_range, , xlYes)
    query_update_table.Name = "JiraQueryUpdateTable"
    query_update_table.TableStyle = "TableStyleLight9"
    'Avoid run-time error if attempted on empty table
    If Not query_update_table.DataBodyRange Is Nothing Then
        query_update_table.DataBodyRange.Interior.ColorIndex = xlNone
    End If
    PopulateDropDown
End Sub

Sub PopulateDropDown()
    'Enable drop-down lists
    With ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range(CELL_ISSUE_TYPE_DROPDOWN).Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:="=Indirect(""JiraIssueTypeIDs[Issue Type]"")"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    With ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range(CELL_STATUS_DROPDOWN).Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:="=Indirect(""JiraStatusIDs[Status]"")"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    With ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range("B3").Validation
        On Error Resume Next 'To avoid runtime error if data validation already present
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:="=Indirect(""Proj1Users[UserID]"")"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    With ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range(CELL_UserID_DROPDOWN).Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:="=Indirect(""Proj1Users[UserID]"")"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    With ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range(CELL_FIX_VERSION_DROPDOWN).Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:="=Indirect(""Proj2[Field Name]"")"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = False
    End With
    With ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range(CELL_PRIORITY_DROPDOWN).Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:="=Indirect(""JiraPriorityIds[Priority]"")"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    With ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range(CELL_OS_DROPDOWN).Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:="=Indirect(""Proj2[Field Name]"")"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = False
    End With
    With ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range(CELL_PLATFORM_DROPDOWN).Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:="=Indirect(""Proj2[Field Name]"")"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = False
    End With
    With ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range(CELL_COMPONENT_DROPDOWN).Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:="=Indirect(""Proj2Components[Component Name]"")"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = False
    End With
    With ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range(CELL_TREND_DROPDOWN).Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:="On Track, Ahead of Commit, Missing Commit, Unknown, No Effort Pending, Non-open Entry"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    With ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range(CELL_ASSESSMENT_DROPDOWN).Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:="Right Estimate, Aggressive Estimate, Conservative Estimate"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
End Sub

'Show progress with provided %
Sub RefreshProgressBar(percent_completion As Integer)
    DoEvents
    ProgressBar.ProgressForeground.Width = (percent_completion / 100) * ProgressBar.ProgressBackground.Width
    ProgressBar.ProgressText.Caption = percent_completion & "% Done"
End Sub

'Verify if the specified field has valid value. This is required to ensure we avoid runtime errors
Function FieldHasValue(json_object As Object, field_name As String) As Boolean
    FieldHasValue = False
    Dim field_value As String
    Dim field_name_array() As String
    Dim jira_json_local As Object
    field_name_array = Split(field_name, "/")
    
    On Error GoTo return_from_function
        If UBound(field_name_array) = 0 Then
            field_value = CallByName(json_object, field_name, VbGet) 'Single level
        Else
            For counter = 0 To UBound(field_name_array) - 1 'Multi level
                Set jira_json_local = CallByName(json_object, field_name_array(counter), VbGet)
            Next
            field_value = CallByName(jira_json_local, field_name_array(UBound(field_name_array)), VbGet)
        End If
        
        FieldHasValue = True
return_from_function:
End Function

'Query jira for name of provided epic key
Function GetEpicName(epic_key As String) As String
    Dim epic_name_local As String
    epic_name_local = ""
    GetHttpRequest (JIRA_API_ISSUE_URL & epic_key & "?fields=" & EXTERNAL_ISSUE_ID_FIELD & "," & EPIC_NAME_FIELD)
    Set jira_json = CallByName(jira_json, "fields", VbGet) 'Get fields object from jira response
    If FieldHasValue(jira_json, EXTERNAL_ISSUE_ID_FIELD) Then
        GetEpicName = "[" & CallByName(jira_json, EXTERNAL_ISSUE_ID_FIELD, VbGet) & "] " & CallByName(jira_json, EPIC_NAME_FIELD, VbGet) & epic_name_local
    Else: GetEpicName = CallByName(jira_json, EPIC_NAME_FIELD, VbGet) & epic_name_local
    End If
End Function

'JIRA PUT/POST requests
Function SendHttpRequest(api_type As String, jira_id As String, update_string As String) As Integer
    If SESSION_ID = "" Then
        SESSION_ID = GetSessionId()
    End If
    
    'Initialize to allow parsing of response text
    Set scriptControl = CreateObject("MSScriptControl.ScriptControl")
    scriptControl.Language = "JScript"
    
    With JiraService
    .Open api_type, JIRA_API_ISSUE_URL & jira_id, False
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "Accept", "application/json"
        .setRequestHeader "Set-Cookie", SESSION_ID
        .setRequestHeader "X-Atlassian-Token", "nocheck"
        On Error Resume Next
        .send update_string
        jira_response = .responseText 'To get the created issue key
        return_value = .status
        If .status = 201 Or .status = 1223 Or .status = 200 Or .status = 204 Or .responseText = "" Then
            'Nothing, just return
        ElseIf .statusText = "Unauthorized" Then
            'Get session ID from JIRA again
            ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range("D1").Value = ""
            SESSION_ID = GetSessionId()
            MsgBox "Authentication successful. Please re-try your action"
            End
        ElseIf .status <> 201 Then
            'Show details for failure
            Set jira_json = .responseBody
            MsgBox "Jira Response: " & .status & " " & .statusText & Chr(10) & Chr(10) & jira_response, vbOKOnly, "Jira Update Failed"
        End If
        .abort
    End With
    SendHttpRequest = return_value
End Function

'JIRA GET requests
Sub GetHttpRequest(search_query As String)
    If SESSION_ID = "" Then
        SESSION_ID = GetSessionId()
    End If
    
    'Initialize to allow parsing of response text
    Set scriptControl = CreateObject("MSScriptControl.ScriptControl")
    scriptControl.Language = "JScript"
    
    With JiraService
    .Open "GET", search_query, False
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "Accept", "application/json"
        .setRequestHeader "X-Atlassian-Token", "nocheck"
        .setRequestHeader "Set-Cookie", SESSION_ID
        .send
        Select Case .status
            Case 200 'Success :)
                response_text = .responseText
                Set jira_json = scriptControl.Eval("(" + .responseText + ")")
            Case 401 '"Unauthorized" 'Get session ID from JIRA again
                ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range("D1").Value = ""
                SESSION_ID = GetSessionId()
                MsgBox "Authentication successful. Please re-try your action"
                End
            Case Else 'Show details of failure
                MsgBox "Jira Response: " & .status & " " & .statusText, vbOKOnly, "JIRA query failed"
                End
        End Select
        .abort
    End With
End Sub

'Return issue type of specified issue key
Function GetIssueType(issue_id As String) As String
    Dim args As String
    args = "&fields=issuetype"
    filter_search_query = JIRA_API_SEARCH_URL & "issuekey in (" & issue_id & ")"
    Call GetHttpRequest(filter_search_query & args)
    GetIssueType = CallByName(CallByName(CallByName(CallByName(CallByName(jira_json, "issues", VbGet), 0, VbGet), "fields", VbGet), "issuetype", VbGet), "name", VbGet)
End Function

'Get details from JIRA based on the selected query
Sub GetJiraJson(query_type As String)
    Initialize 'Mandatory
    Dim args As String
    args = "&fields=issuetype,summary,assignee,fixVersions,priority,duedate,components,issuelinks,status,timetracking,"
    args = args & START_DATE_FIELD & "," & END_DATE_FIELD & "," & CUSTOM_FIELD_1 & "," & CUSTOM_FIELD_2 & "," & EPIC_LINK_FIELD & "&maxResults=-1" 'Max of 1000 results
    
    Dim filter_search_query As String
    Select Case query_type
        Case "filter"
            GetHttpRequest (JIRA_API_FILTER_URL & ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range("B2").Value & "?expand")
            filter_search_query = CallByName(jira_json, "searchUrl", VbGet)
        Case "assignee"
            filter_search_query = JIRA_API_SEARCH_URL & "assignee in (" & ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range("B3").Value & ")"
        Case "epic"
            filter_search_query = JIRA_API_SEARCH_URL & "%22Epic Link%22 in (" & ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range("B5").Value & ")"
        Case "id"
            filter_search_query = JIRA_API_SEARCH_URL & "issuekey in (" & ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range("B4").Value & ")"
        Case "text"
            filter_search_query = JIRA_API_SEARCH_URL & "text ~ %22" & ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range("B1").Value & "%22"
        Case Else
            'Do nothing
    End Select
    
    'Status
    'If ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range("C2").Value <> "" Then
    '    filter_search_query = filter_search_query & " AND status in (" & ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range("B6").Value & ")"
    'End If
    
    Call GetHttpRequest(filter_search_query & args)
    
    'Exit if no issues linked to epic
    If query_type = "epic" And CallByName(jira_json, "total", VbGet) = 0 Then
        MsgBox "No linked issues found."
        End
    End If
    
    Dim total_issues As Long
    'Total issues vary based on the count of issues in the filter
    total_issues = WorksheetFunction.Min(CallByName(jira_json, "maxResults", VbGet), CallByName(jira_json, "total", VbGet))
    If total_issues = 0 Then
        MsgBox "No issues found", vbOKOnly, "Empty data from query"
        End
    Else: ReDim jira_entry(CallByName(jira_json, "total", VbGet) - 1, QUERY_UPDATE_COLUMN_COUNT)
    End If
    
    'Parse JSON details of each issue
    For Count = 0 To total_issues - 1
        Set jira_json_fields = CallByName(CallByName(CallByName(jira_json, "issues", VbGet), Count, VbGet), "fields", VbGet)
        Set jira_issue_json = CallByName(CallByName(jira_json, "issues", VbGet), Count, VbGet)
        GetJiraDetails (Count) 'Extract required details from JSON
    Next
    
    'Logic to get more than 1000 issues
    Dim loop_count As Integer
    loop_count = 1
    While (CallByName(jira_json, "total", VbGet) - (1000 * loop_count)) / (1000) > 1
        GetHttpRequest (filter_search_query & args & "&startAt=" & (1000 * loop_count))
        total_issues = WorksheetFunction.Min(CallByName(jira_json, "maxResults", VbGet), CallByName(jira_json, "total", VbGet))
        For Count = 0 To total_issues - 1
            Set jira_json_fields = CallByName(CallByName(CallByName(jira_json, "issues", VbGet), Count, VbGet), "fields", VbGet)
            Set jira_issue_json = CallByName(CallByName(jira_json, "issues", VbGet), Count, VbGet)
            GetJiraDetails (Count + (1000 * loop_count))
        Next
        loop_count = loop_count + 1
    Wend
    
    'Logic to get issues that are between x001 and x999 (beyond the first 1000)
    If CallByName(jira_json, "total", VbGet) Mod (1000 * loop_count) > 0 And CallByName(jira_json, "total", VbGet) / (1000 * loop_count) > 1 Then
        GetHttpRequest (filter_search_query & args & "&startAt=" & (1000 * loop_count))
        total_issues = CallByName(jira_json, "total", VbGet) Mod (1000 * loop_count)
        For Count = 0 To total_issues - 1
            Set jira_json_fields = CallByName(CallByName(CallByName(jira_json, "issues", VbGet), Count, VbGet), "fields", VbGet)
            Set jira_issue_json = CallByName(CallByName(jira_json, "issues", VbGet), Count, VbGet)
            GetJiraDetails (Count + (1000 * loop_count))
        Next
    End If
    
    'Write queried data to sheet
    'query_update_table.Resize query_update_table.Range.Resize(CallByName(jira_json, "total", VbGet) + 1)
    ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range(QUERY_UPDATE_WRITE_COL & (CallByName(jira_json, "total", VbGet) + QUERY_UPDATE_ROW_OFFSET)).Value = jira_entry()
    query_update_table.Resize query_update_table.Range.Resize(CallByName(jira_json, "total", VbGet) + 1)
    ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range(CELL_QUERY_UPDATE_DATABODYRANGE_START).Select
    ThisWorkbook.RefreshAll
End Sub

'Query JIRA with assignee
Sub GetIssuesByAssignee()
    start_time = DateTime.Now
    GetJiraJson ("assignee")
    Cleanup
End Sub

'Query JIRA with filter ID
Sub GetIssuesByFilter()
    start_time = DateTime.Now
    GetJiraJson ("filter")
    Cleanup
End Sub

'Query JIRA for issues in epic
Sub GetEpicIssues()
    start_time = DateTime.Now
    
    'Ensure entered ID is of an epic
    If GetIssueType(ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range("B5").Value) = "Epic" Then
        GetJiraJson ("epic")
    Else
        MsgBox ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range("B5").Value & " is not an epic", vbOKOnly, "Invalid epic key"
    End If
    Cleanup
End Sub

'Query JIRA with issue ID
Sub GetIssueById()
    start_time = DateTime.Now
    GetJiraJson ("id")
    Cleanup
End Sub

'Query JIRA with text search
Sub GetIssueByText()
    start_time = DateTime.Now
    GetJiraJson ("text")
    Cleanup
End Sub

'Parse JSON from jira to extract required details
Sub GetJiraDetails(row As Integer)
    On Error Resume Next
    
    'Few multi-select fields do not have a parameter to tell the entry count in the respective JIRA
    'Here the length of each entry is 16, so the logic below is to derive the value count for these fields
    If Len(jira_json_fields.issuelinks) > 0 Then
        issue_links_count = 1 + (Len(jira_json_fields.issuelinks) - 15) / 16
    End If
    If Len(jira_json_fields.components) > 0 Then
        custom_field_2_count = 1 + (Len(jira_json_fields.components) - 15) / 16
    End If
    If Len(jira_json_fields.customfield_10700) > 0 Then
        custom_field_1_count = 1 + (Len(jira_json_fields.customfield_10700) - 15) / 16
    End If
    If Len(jira_json_fields.fixVersions) > 0 Then
        fix_version_count = 1 + (Len(jira_json_fields.fixVersions) - 15) / 16
    End If
    
    jira_entry(row, ISSUE_KEY_COLUMN) = CallByName(jira_issue_json, "key", VbGet)
    jira_entry(row, ISSUE_TYPE_COLUMN) = CallByName(CallByName(jira_json_fields, "issuetype", VbGet), "name", VbGet)
    jira_entry(row, ISSUE_STATUS_COLUMN) = CallByName(CallByName(jira_json_fields, "status", VbGet), "name", VbGet)
    jira_entry(row, ISSUE_SUMMARY_COLUMN) = CallByName(jira_json_fields, "summary", VbGet)
    jira_entry(row, ISSUE_ASSIGNEE_COLUMN) = CallByName(CallByName(jira_json_fields, "assignee", VbGet), "name", VbGet)
    jira_entry(row, ISSUE_EPIC_LINK_COLUMN) = CallByName(jira_json_fields, EPIC_LINK_FIELD, VbGet)
    
    For Each issue_link In CallByName(jira_json_fields, "issuelinks", VbGet)
        If CallByName(CallByName(issue_link, "type", VbGet), "name", VbGet) = "Blocks" Then
            If jira_entry(row, ISSUE_BLOCKED_BY_COLUMN) = "" Then
                jira_entry(row, ISSUE_BLOCKED_BY_COLUMN) = CallByName(CallByName(issue_link, "inwardIssue", VbGet), "key", VbGet)
            Else
                jira_entry(row, ISSUE_BLOCKED_BY_COLUMN) = jira_entry(row, ISSUE_BLOCKED_BY_COLUMN) & ", " & CallByName(CallByName(issue_link, "inwardIssue", VbGet), "key", VbGet)
            End If
        End If
    Next
    
    For Each issue_link In CallByName(jira_json_fields, "issuelinks", VbGet)
        If CallByName(CallByName(issue_link, "type", VbGet), "name", VbGet) = "Blocks" Then
            If jira_entry(row, ISSUE_BLOCKS_COLUMN) = "" Then
                jira_entry(row, ISSUE_BLOCKS_COLUMN) = CallByName(CallByName(issue_link, "outwardIssue", VbGet), "key", VbGet)
            Else
                jira_entry(row, ISSUE_BLOCKS_COLUMN) = jira_entry(row, ISSUE_BLOCKS_COLUMN) & ", " & CallByName(CallByName(issue_link, "outwardIssue", VbGet), "key", VbGet)
            End If
        End If
    Next
    
    If Len(jira_json_fields.fixVersions) > 0 Then
        jira_entry(row, ISSUE_FIX_VERSION_COLUMN) = CallByName(CallByName(CallByName(jira_json_fields, "fixVersions", VbGet), 0, VbGet), "name", VbGet)
        For Count = 1 To fix_version_count - 1
            jira_entry(row, ISSUE_FIX_VERSION_COLUMN) = jira_entry(row, ISSUE_FIX_VERSION_COLUMN) + ", " + CallByName(CallByName(CallByName(jira_json_fields, "fixVersions", VbGet), Count, VbGet), "name", VbGet)
        Next
    End If
    jira_entry(row, ISSUE_PRIORITY_COLUMN) = CallByName(CallByName(jira_json_fields, "priority", VbGet), "name", VbGet)
    If Len(jira_json_fields.CUSTOM_FIELD_2) > 0 Then
        jira_entry(row, ISSUE_CUSTOM_FIELD_2_COLUMN) = CallByName(CallByName(CallByName(jira_json_fields, CUSTOM_FIELD_2, VbGet), 0, VbGet), "value", VbGet)
        For Count = 1 To custom_field_1_count - 1
            jira_entry(row, ISSUE_CUSTOM_FIELD_2_COLUMN) = jira_entry(row, ISSUE_CUSTOM_FIELD_2_COLUMN) + ", " + CallByName(CallByName(CallByName(jira_json_fields, CUSTOM_FIELD_2, VbGet), Count, VbGet), "value", VbGet)
        Next
    End If
    jira_entry(row, ISSUE_DUE_DATE_COLUMN) = CallByName(jira_json_fields, "duedate", VbGet)
    jira_entry(row, ISSUE_START_DATE_COLUMN) = Left(CallByName(jira_json_fields, START_DATE_FIELD, VbGet), 10) 'Get only date section from ISO format
    jira_entry(row, ISSUE_END_DATE_COLUMN) = Left(CallByName(jira_json_fields, END_DATE_FIELD, VbGet), 10) 'Get only date section from ISO format
    If Len(jira_json_fields.CUSTOM_FIELD_1) > 0 Then
        jira_entry(row, ISSUE_CUSTOM_FIELD_1_COLUMN) = CallByName(CallByName(CallByName(jira_json_fields, CUSTOM_FIELD_1, VbGet), 0, VbGet), "value", VbGet)
        For Count = 1 To custom_field_1_count - 1
            jira_entry(row, ISSUE_CUSTOM_FIELD_1_COLUMN) = jira_entry(row, ISSUE_CUSTOM_FIELD_1_COLUMN) + ", " + CallByName(CallByName(CallByName(jira_json_fields, CUSTOM_FIELD_1, VbGet), Count, VbGet), "value", VbGet)
        Next
    End If
    If Len(jira_json_fields.components) > 0 Then
        jira_entry(row, ISSUE_COMPONENTS_COLUMN) = CallByName(CallByName(CallByName(jira_json_fields, "components", VbGet), 0, VbGet), "name", VbGet)
        For Count = 1 To custom_field_2_count - 1
            jira_entry(row, ISSUE_COMPONENTS_COLUMN) = jira_entry(row, ISSUE_COMPONENTS_COLUMN) + ", " + CallByName(CallByName(CallByName(jira_json_fields, "components", VbGet), Count, VbGet), "name", VbGet)
        Next
    End If
    
    If FieldHasValue(jira_json_fields, "timetracking/originalEstimate") Then
        jira_entry(row, ISSUE_ORIGINAL_ESTIMATE_COLUMN) = CallByName(CallByName(jira_json_fields, "timetracking", VbGet), "originalEstimateSeconds", VbGet) / 3600
    Else
        jira_entry(row, ISSUE_ORIGINAL_ESTIMATE_COLUMN) = "0"
    End If
    
    If FieldHasValue(jira_json_fields, "timetracking/remainingEstimate") Then
        jira_entry(row, ISSUE_REMAINING_ESTIMATE_COLUMN) = CallByName(CallByName(jira_json_fields, "timetracking", VbGet), "remainingEstimateSeconds", VbGet) / 3600
    Else
        jira_entry(row, ISSUE_REMAINING_ESTIMATE_COLUMN) = "0"
    End If
    
    If FieldHasValue(jira_json_fields, "timetracking/timeSpent") Then
        jira_entry(row, ISSUE_TIME_SPENT_COLUMN) = CallByName(CallByName(jira_json_fields, "timetracking", VbGet), "timeSpentSeconds", VbGet) / 3600
    Else
        jira_entry(row, ISSUE_TIME_SPENT_COLUMN) = "0"
    End If

    If jira_entry(row, ISSUE_ORIGINAL_ESTIMATE_COLUMN) <> "0" Then
        jira_entry(row, ISSUE_BANDWIDTH_COLUMN) = jira_entry(row, ISSUE_ORIGINAL_ESTIMATE_COLUMN) / GetWorkHours(jira_entry(row, ISSUE_START_DATE_COLUMN), jira_entry(row, ISSUE_END_DATE_COLUMN))
    Else: jira_entry(row, ISSUE_BANDWIDTH_COLUMN) = "0%"
    End If
    
    If jira_entry(row, ISSUE_REMAINING_ESTIMATE_COLUMN) <> "0" And dict_open_states.Exists(jira_entry(row, ISSUE_STATUS_COLUMN)) Then
        jira_entry(row, ISSUE_BEST_END_COLUMN) = WorksheetFunction.WorkDay(DateTime.Date, (jira_entry(row, ISSUE_REMAINING_ESTIMATE_COLUMN) / WORK_HOURS_PER_DAY), holidays.DataBodyRange)
    End If
    
    Dim pending_effort_days As Double
    If dict_open_states.Exists(jira_entry(row, ISSUE_STATUS_COLUMN)) Then
        If jira_entry(row, ISSUE_REMAINING_ESTIMATE_COLUMN) = "0" Then
            jira_entry(row, ISSUE_TREND_COLUMN) = "No Effort Pending"
        ElseIf jira_entry(row, ISSUE_BANDWIDTH_COLUMN) = "" Then
            jira_entry(row, ISSUE_TREND_COLUMN) = "Unknown"
        Else
            pending_effort_days = (jira_entry(row, ISSUE_REMAINING_ESTIMATE_COLUMN) / jira_entry(row, ISSUE_BANDWIDTH_COLUMN)) / WORK_HOURS_PER_DAY
            If WorksheetFunction.WorkDay(DateTime.Date, pending_effort_days, holidays.DataBodyRange) > CDate(jira_entry(row, ISSUE_END_DATE_COLUMN)) Then
                jira_entry(row, ISSUE_TREND_COLUMN) = "Missing Commit"
            ElseIf WorksheetFunction.WorkDay(DateTime.Date, pending_effort_days, holidays.DataBodyRange) < CDate(jira_entry(row, ISSUE_END_DATE_COLUMN)) Then
                jira_entry(row, ISSUE_TREND_COLUMN) = "Ahead of Commit"
            ElseIf WorksheetFunction.WorkDay(DateTime.Date, pending_effort_days, holidays.DataBodyRange) = CDate(jira_entry(row, ISSUE_END_DATE_COLUMN)) Then
                jira_entry(row, ISSUE_TREND_COLUMN) = "On Track"
            End If
        End If
    Else:
        jira_entry(row, ISSUE_TREND_COLUMN) = "Non-open Entry"
        If CDbl(jira_entry(row, ISSUE_REMAINING_ESTIMATE_COLUMN)) = 0 Then
            jira_entry(row, ISSUE_ASSESSMENT_COLUMN) = "Right Estimate"
        ElseIf CDbl(jira_entry(row, ISSUE_REMAINING_ESTIMATE_COLUMN)) > 0 Then
            jira_entry(row, ISSUE_ASSESSMENT_COLUMN) = "Conservative Estimate"
        ElseIf CDbl(jira_entry(row, ISSUE_TIME_SPENT_COLUMN)) > CDbl(jira_entry(row, ISSUE_ORIGINAL_ESTIMATE_COLUMN)) Then
            jira_entry(row, ISSUE_ASSESSMENT_COLUMN) = "Aggressive Estimate"
        End If
    End If
End Sub

'Encode user credentials
Private Function UserPassBase64() As String
    
    Dim objXML As MSXML2.DOMDocument60
    Dim objNoe As MSXML2.IXMLDOMElement
    Dim arrData() As Byte
    
    arrData = StrConv("username:password", vbFromUnicode)
    
    Set objXML = New MSXML2.DOMDocument60
    Set objNode = objXML.createElement("b64")
    objNode.DataType = "bin.base64"
    objNode.nodeTypedValue = arrData
    
    UserPassBase64 = objNode.Text
    Dim encoded_credentials As String
    encoded_credentials = objNode.Text

End Function

'Get session ID from sheet
Function GetSessionId() As String
    If ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range("D1").Value = "" _
        Or DateDiff("n", ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range("B7").Value, DateTime.Now) > 30 Then 'Key not used for more than 30mins
        SessionIdQuery.Show 'Get details from user
    End If
    
    GetSessionId = ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range("D1").Value
End Function

'Delete session ID
Function DeleteSessionID()
    Dim JiraAuth As New MSXML2.XMLHTTP60
    With JiraAuth
        .Open "DELETE", JIRA_API_AUTH_URL, False
        .send
    End With
End Function

Sub EnableTracking()
    track_changes = True
    Application.EnableEvents = True
End Sub

Function GetWorkHours(issue_start As Variant, issue_end As Variant) As Double
    Dim work_hours As Double
    work_hours = WorksheetFunction.NetworkDays(issue_start, issue_end, holidays.DataBodyRange) * WORK_HOURS_PER_DAY
    GetWorkHours = work_hours
End Function
