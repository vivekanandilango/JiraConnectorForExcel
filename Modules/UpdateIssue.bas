Attribute VB_Name = "UpdateIssue"
'Used across functions in UpdateIssue module
Dim issues_to_update() As String
Dim progress_bar_counter As Integer

Function CellModified(cell_address As String) As Boolean
    Dim cell_modified As Boolean
    cell_modified = False
    If ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range(cell_address).Interior.ColorIndex = 6 Then
        cell_modified = True
    End If
    CellModified = cell_modified
End Function

Sub AddCommentOfSelectedRow(comment_to_add As String)
    comment_to_add = Replace(comment_to_add, Chr(10), "\n")
    Dim jira_key As String
    jira_key = ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range("A" & ActiveCell.row).Value
    Call SendHttpRequest(API_PUT, jira_key, "{""update"": {""comment"": [ { ""add"": {""body"": """ & comment_to_add & """} }]}}")
    Cleanup
End Sub

Sub ChangeSummaryOfSelectedRow()
    Dim summary_to_update As String
    Dim jira_id As String
    jira_id = ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range("A" & ActiveCell.row).Value
    summary_to_update = ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range("D" & ActiveCell.row).Value
    Call SendHttpRequest(API_PUT, jira_id, "{""fields"":{""summary"":""" & summary_to_update & """}}")
    Cleanup
End Sub

Sub ChangeStatusOfSelectedRow()
    Dim status As String
    status = ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range("C" & ActiveCell.row).Value
    transition_id = GetTransitionId(jira_id, status)
    Call SendHttpRequest(API_POST, jira_id & "/transitions?expand=transitions.fields", "{""update"": {""comment"": [ { ""add"": {""body"": ""Status change to " & status & """} }]},""transition"":{""id"":""" & transition_id & """}}")
    Cleanup
End Sub

Sub ChangeAssigneeOfSelectedRow()
    Dim ASSIGNEE As String
    ASSIGNEE = ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range("E" & ActiveCell.row).Value
    Call SendHttpRequest(API_PUT, jira_id, "{""fields"":{""assignee"":{""name"":""" & ASSIGNEE & """}}}")
    Cleanup
End Sub

'Update priority if changed
Sub ChangePriorityOfSelectedRow()
    PRIORITY = ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range("G" & ActiveCell.row).Value
    jira_key = ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range("A" & ActiveCell.row).Value
    If CellModified("G" & ActiveCell.row) Then
        return_value = SendHttpRequest(API_PUT, jira_id, "{""update"":{""priority"":[{""set"":{""name"" : """ & PRIORITY & """}}]}}")
        If return_value = 201 Or return_value = 1223 Then
            ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range("G" & ActiveCell.row).Interior.ColorIndex = 4
        End If
    Else: MsgBox "Cell value not changed"
    End If
    Cleanup
End Sub

Sub ChangeFixVersionOfSelectedRow()
    Dim FIX_VERSION As String
    FIX_VERSION = ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range("F" & ActiveCell.row).Value
    jira_key = ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range("A" & ActiveCell.row).Value
    Call SendHttpRequest(API_PUT, jira_id, "{""fields"":{""fixVersions"":[" & GetMultiValueString("name", FIX_VERSION) & "]}}")
    Cleanup
End Sub

Sub ChangeCustomField1OfSelectedRow()
    Dim CustomField1 As String
    CustomField1 = ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range("L" & ActiveCell.row).Value
    jira_key = ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range("A" & ActiveCell.row).Value
    Call SendHttpRequest(API_PUT, jira_id, "{""fields"":{""" & CUSTOM_FIELD_1 & """:[" & GetMultiValueString("value", CustomField1) & "]}}")
End Sub

Sub ChangeCustomField2OfSelectedRow()
    Dim CustomField2 As String
    CustomField2 = ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range("H" & ActiveCell.row).Value
    jira_key = ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range("A" & ActiveCell.row).Value
    Call SendHttpRequest(API_PUT, jira_id, "{""fields"":{""" & CUSTOM_FIELD_2 & """:[" & GetMultiValueString("value", CustomField2) & "]}}")
End Sub

Sub ChangeComponentOfSelectedRow()
    Dim COMPONENT As String
    COMPONENT = ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range("M" & ActiveCell.row).Value
    jira_key = ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range("A" & ActiveCell.row).Value
    Call SendHttpRequest(API_PUT, jira_id, "{""fields"":{""components"":[" & GetMultiValueString("name", COMPONENT) & "]}}")
    Cleanup
End Sub

Sub ChangeDueDateOfSelectedRow()
    Dim DUE_DATE As String
    DUE_DATE = ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range("I" & ActiveCell.row).Value
    jira_key = ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range("A" & ActiveCell.row).Value
    Call SendHttpRequest(API_PUT, jira_id, "{""fields"":{""duedate"":""" & Format(DUE_DATE, "yyyy-mm-dd") & """}}")
    Cleanup
End Sub

Sub ChangeStartDateOfSelectedRow()
    Dim START_DATE As String
    START_DATE = ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range("J" & ActiveCell.row).Value
    jira_key = ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range("A" & ActiveCell.row).Value
    Call SendHttpRequest(API_PUT, jira_id, "{""fields"":{""customfield_xxxxx"":""" & Format(START_DATE, "yyyy-mm-dd") & "T00:00:00.000+0530""}}")
    Cleanup
End Sub

Sub ChangeEndDateOfSelectedRow()
    Dim END_DATE As String
    END_DATE = ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range("K" & ActiveCell.row).Value
    jira_key = ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range("A" & ActiveCell.row).Value
    Call SendHttpRequest(API_PUT, jira_id, "{""fields"":{""customfield_xxxxx"":""" & Format(END_DATE, "yyyy-mm-dd") & "T00:00:00.000+0530""}}")
    Cleanup
End Sub

Sub ChangeOriginalEstimate()
    Dim ORIGINAL_ESTIMATE As String
    ORIGINAL_ESTIMATE = ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range("N" & ActiveCell.row).Value
    jira_key = ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range("A" & ActiveCell.row).Value
    Call SendHttpRequest(API_PUT, jira_id, "{""update"": {""timetracking"": [ { ""edit"": {""originalEstimate"": """ & ORIGINAL_ESTIMATE & "h""} }]}}")
End Sub

Sub ChangeRemainingEstimate()
    Dim remaining_estimate As String
    remaining_estimate = ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range("O" & ActiveCell.row).Value
    jira_key = ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range("A" & ActiveCell.row).Value
    Call SendHttpRequest(API_PUT, jira_id, "{""update"": {""timetracking"": [ { ""edit"": {""remainingEstimate"": """ & remaining_estimate & "h""} }]}}")
End Sub

'Creates JSON to be sent to JIRA
Sub BuildIssueUpdateJson(issue_field() As String, update_type As String)
    Dim DUE_DATE As String
    Dim START_DATE As String
    Dim END_DATE As String
    Dim ASSIGNEE As String
    Dim SUMMARY As String
    Dim CUSTOM_FIELD_1 As String
    Dim CUSTOM_FIELD_2 As String
    Dim COMPONENT As String
    Dim FIX_VERSION As String
    Dim EPIC_LINK As String
    Dim BLOCKS As String
    Dim BLOCKED_BY As String
    
    Dim PRIORITY As String
    Dim ORIGINAL_ESTIMATE As String
    Dim description As String
    Dim COMMENT As String
    
    Dim first_field As Boolean
    Dim separator As String
    separator = ","
    
    Select Case update_type
        Case "worklog"
            UPDATE_ISSUE_JSON = "{""comment"": """ & issue_field(ISSUE_COMMENTS_COLUMN) & _
                                """,""timeSpent"": " & issue_field(ISSUE_TIME_SPENT_COLUMN) & "}"
            
        Case "status"
            transition_id = GetTransitionId(issue_field(ISSUE_KEY_COLUMN), issue_field(ISSUE_STATUS_COLUMN))
            UPDATE_ISSUE_JSON = "{""update"": {""comment"": [ { ""add"": {""body"": ""Status change to " & issue_field(ISSUE_STATUS_COLUMN) & _
                                """} }]},""transition"":{""id"":""" & transition_id & """}}"
                                
        Case "estimate"
            UPDATE_ISSUE_JSON = "{""update"": {""timetracking"": [ { ""edit"": {""remainingEstimate"": """ & issue_field(ISSUE_REMAINING_ESTIMATE_COLUMN) & "h""} }]}}"
            
        Case "field"
            first_field = True
            UPDATE_ISSUE_JSON = """fields"":{"
            If issue_field(ISSUE_DUE_DATE_COLUMN) <> "" Then
                DUE_DATE = """duedate"":""" & Format(issue_field(ISSUE_DUE_DATE_COLUMN), "yyyy-mm-dd") & """"
                UPDATE_ISSUE_JSON = UPDATE_ISSUE_JSON & DUE_DATE
                first_field = False
            End If
            If issue_field(ISSUE_START_DATE_COLUMN) <> "" Then
                START_DATE = """" & START_DATE_FIELD & """:""" & Format(issue_field(ISSUE_START_DATE_COLUMN), "yyyy-mm-dd") & "T00:00:00.000+0530"""
                If first_field = True Then
                    UPDATE_ISSUE_JSON = UPDATE_ISSUE_JSON & START_DATE
                    first_field = False
                Else: UPDATE_ISSUE_JSON = UPDATE_ISSUE_JSON & separator & START_DATE
                End If
            End If
            If issue_field(ISSUE_END_DATE_COLUMN) <> "" Then
                END_DATE = """" & END_DATE_FIELD & """:""" & Format(issue_field(ISSUE_END_DATE_COLUMN), "yyyy-mm-dd") & "T00:00:00.000+0530"""
                If first_field = True Then
                    UPDATE_ISSUE_JSON = UPDATE_ISSUE_JSON & END_DATE
                    first_field = False
                Else: UPDATE_ISSUE_JSON = UPDATE_ISSUE_JSON & separator & END_DATE
                End If
            End If
            If issue_field(ISSUE_SUMMARY_COLUMN) <> "" Then
                SUMMARY = """summary"":""" & issue_field(ISSUE_SUMMARY_COLUMN) & """"
                If first_field = True Then
                    UPDATE_ISSUE_JSON = UPDATE_ISSUE_JSON & SUMMARY
                    first_field = False
                Else: UPDATE_ISSUE_JSON = UPDATE_ISSUE_JSON & separator & SUMMARY
                End If
            End If
            If issue_field(ISSUE_ASSIGNEE_COLUMN) <> "" Then
                ASSIGNEE = """assignee"":{""name"":""" & issue_field(ISSUE_ASSIGNEE_COLUMN) & """}"
                If first_field = True Then
                    UPDATE_ISSUE_JSON = UPDATE_ISSUE_JSON & ASSIGNEE
                    first_field = False
                Else: UPDATE_ISSUE_JSON = UPDATE_ISSUE_JSON & separator & ASSIGNEE
                End If
            End If
            If issue_field(ISSUE_CUSTOM_FIELD_1_COLUMN) <> "" Then
                CUSTOM_FIELD_1 = """" & CUSTOM_FIELD_1 & """:[" & GetMultiValueString("value", issue_field(ISSUE_CUSTOM_FIELD_1_COLUMN)) & "]"
                If first_field = True Then
                    UPDATE_ISSUE_JSON = UPDATE_ISSUE_JSON & CUSTOM_FIELD_1
                    first_field = False
                Else: UPDATE_ISSUE_JSON = UPDATE_ISSUE_JSON & separator & CUSTOM_FIELD_1
                End If
            End If
            If issue_field(ISSUE_CUSTOM_FIELD_2_COLUMN) <> "" Then
               CUSTOM_FIELD_2 = """" & CUSTOM_FIELD_2 & """:[" & GetMultiValueString("value", issue_field(ISSUE_CUSTOM_FIELD_2_COLUMN)) & "]"
                If first_field = True Then
                    UPDATE_ISSUE_JSON = UPDATE_ISSUE_JSON & CUSTOM_FIELD_2
                    first_field = False
                Else: UPDATE_ISSUE_JSON = UPDATE_ISSUE_JSON & separator & CUSTOM_FIELD_2
                End If
            End If
            If issue_field(ISSUE_COMPONENTS_COLUMN) <> "" Then
                COMPONENT = """components"":[" & GetMultiValueString("name", issue_field(ISSUE_COMPONENTS_COLUMN)) & "]"
                If first_field = True Then
                    UPDATE_ISSUE_JSON = UPDATE_ISSUE_JSON & COMPONENT
                    first_field = False
                Else: UPDATE_ISSUE_JSON = UPDATE_ISSUE_JSON & separator & COMPONENT
                End If
            End If
            If issue_field(ISSUE_FIX_VERSION_COLUMN) <> "" Then
                FIX_VERSION = """fixVersions"":[" & GetMultiValueString("name", issue_field(ISSUE_FIX_VERSION_COLUMN)) & "]"
                If first_field = True Then
                    UPDATE_ISSUE_JSON = UPDATE_ISSUE_JSON & FIX_VERSION
                    first_field = False
                Else: UPDATE_ISSUE_JSON = UPDATE_ISSUE_JSON & separator & FIX_VERSION
                End If
            End If
            If issue_field(ISSUE_EPIC_LINK_COLUMN) <> "" Then
                EPIC_LINK = """" & BLOCKS_FIELD & """:[{""set"":""" & issue_field(ISSUE_BLOCKS_COLUMN) & """}]"
                If first_field = True Then
                    UPDATE_ISSUE_JSON = UPDATE_ISSUE_JSON & EPIC_LINK
                    first_field = False
                Else: UPDATE_ISSUE_JSON = UPDATE_ISSUE_JSON & separator & EPIC_LINK
                End If
            End If
            UPDATE_ISSUE_JSON = UPDATE_ISSUE_JSON & "}"
            
        Case "update"
            first_field = True
            UPDATE_ISSUE_JSON = """update"":{"
            If issue_field(ISSUE_PRIORITY_COLUMN) <> "" Then
                PRIORITY = """priority"":[{""set"":{""name"" : """ & issue_field(ISSUE_PRIORITY_COLUMN) & """}}]"
                If first_field = True Then
                    UPDATE_ISSUE_JSON = UPDATE_ISSUE_JSON & PRIORITY
                    first_field = False
                Else: UPDATE_ISSUE_JSON = UPDATE_ISSUE_JSON & separator & PRIORITY
                End If
            End If
            If issue_field(ISSUE_ORIGINAL_ESTIMATE_COLUMN) <> "" Then
                ORIGINAL_ESTIMATE = """timetracking"":[{""edit"":{""originalEstimate"":""" & issue_field(ISSUE_ORIGINAL_ESTIMATE_COLUMN) & "h""}}]"
                If first_field = True Then
                    UPDATE_ISSUE_JSON = UPDATE_ISSUE_JSON & ORIGINAL_ESTIMATE
                    first_field = False
                Else: UPDATE_ISSUE_JSON = UPDATE_ISSUE_JSON & separator & ORIGINAL_ESTIMATE
                End If
            End If
            If issue_field(ISSUE_COMMENTS_COLUMN) <> "" Then
                COMMENT = """comment"":[{""add"":{""body"":""" & Replace(issue_field(ISSUE_COMMENTS_COLUMN), Chr(10), "\n") & """}}]"
                If first_field = True Then
                    UPDATE_ISSUE_JSON = UPDATE_ISSUE_JSON & COMMENT
                    first_field = False
                Else: UPDATE_ISSUE_JSON = UPDATE_ISSUE_JSON & separator & COMMENT
                End If
            End If
            If issue_field(ISSUE_BLOCKS_COLUMN) <> "" Then
                BLOCKS = """issuelinks"":[{""add"":{""type"":{""name"":""Blocks"",""outward"":""blocks""},""outwardIssue"":{""key"":""" & issue_field(ISSUE_BLOCKS_COLUMN) & """}}}]"
                If first_field = True Then
                    UPDATE_ISSUE_JSON = UPDATE_ISSUE_JSON & BLOCKS
                    first_field = False
                Else: UPDATE_ISSUE_JSON = UPDATE_ISSUE_JSON & separator & BLOCKS
                End If
            End If
            If issue_field(ISSUE_BLOCKED_BY_COLUMN) <> "" Then
                BLOCKED_BY = """issuelinks"":[{""add"":{""type"":{""name"":""Blocks"",""inward"":""is blocked by""},""inwardIssue"":{""key"":""" & issue_field(ISSUE_BLOCKED_BY_COLUMN) & """}}}]"
                If first_field = True Then
                    UPDATE_ISSUE_JSON = UPDATE_ISSUE_JSON & BLOCKED_BY
                    first_field = False
                Else: UPDATE_ISSUE_JSON = UPDATE_ISSUE_JSON & separator & BLOCKED_BY
                End If
            End If
            UPDATE_ISSUE_JSON = UPDATE_ISSUE_JSON & "}"
            
        Case Else
            'Do nothing
    End Select
End Sub

Sub AddWorkLogOfSelectedRow(worklog_hours As String, worklog_comment As String)
    worklog_comment = Replace(worklog_comment, Chr(10), "\n")
    Dim jira_key As String
    jira_key = ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range("A" & ActiveCell.row).Value
    Call SendHttpRequest(API_POST, jira_key & "/worklog", "{""comment"": """ & worklog_comment & """,""timeSpent"": " & worklog_hours & "}")
    Cleanup
End Sub

Function GetTransitionId(issue_id As String, transition_status As String) As String

    If SESSION_ID = "" Then
        SESSION_ID = GetSessionId()
    End If
    
    Set scriptControl = CreateObject("MSScriptControl.ScriptControl")
    scriptControl.Language = "JScript"
    
    With JiraService
    .Open "GET", JIRA_API_ISSUE_URL & issue_id & "/transitions", False
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "Accept", "application/json"
        .setRequestHeader "X-Atlassian-Token", "nocheck"
        .setRequestHeader "Set-Cookie", SESSION_ID
        .send
        If .status = 401 Then 'Handle access denied error
            ThisWorkbook.Worksheets("Home").Range("D1").Value = ""
            SESSION_ID = GetSessionId()
            .setRequestHeader "Set-Cookie", SESSION_ID
            .send
            If .status = 200 Then 'Successful authentication
                MsgBox "Authentication completed. Please re-try your action", vbOKOnly, "Authentication success"
            End If
        End If
        If .status = 200 Then
            response_string = .responseText
            Set jira_json = scriptControl.Eval("(" + .responseText + ")")
            Set jira_json = CallByName(jira_json, "transitions", VbGet)
        End If
        .abort
    End With
    Dim queried_ids() As Variant
    ReDim queried_ids(CallByName(jira_json, "length", VbGet) - 1, 1)
    For counter = 0 To CallByName(jira_json, "length", VbGet) - 1
        If (CallByName(CallByName(jira_json, counter, VbGet), "name", VbGet)) = transition_status Then
            GetTransitionId = CallByName(CallByName(jira_json, counter, VbGet), "id", VbGet)
            Exit For
        End If
    Next
End Function

'Parse the queried JIRA table for any updates
Sub GetModifiedIssues()
    Set query_update_table = ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).ListObjects("JiraQueryUpdateTable")
    PopulateColumnUpdateType
    
    update_call_count = 0
    Dim field_put_generic As Boolean
    Dim update_put_generic As Boolean
    Dim rem_estimate_put_generic As Boolean
    Dim worklog_post As Boolean
    Dim status_post As Boolean
    
    field_put_generic = False
    update_put_generic = False
    rem_estimate_put_generic = False
    worklog_post = False
    status_post = False
    
    query_row_count = query_update_table.DataBodyRange.Rows.Count
    query_col_count = query_update_table.DataBodyRange.Columns.Count
    
    'Change size based on the available issues
    ReDim issues_to_update(query_row_count - 1, query_col_count + 2) As String
    
    For row_index = 1 To query_row_count
        issues_to_update(row_index - 1, 0) = query_update_table.DataBodyRange(row_index, 1).Value
        Dim update_calls_for_this_issue As Integer
        update_calls_for_this_issue = 0
        For col_index = 1 To query_col_count
            'Cell with yellow background color is considered modified
            If query_update_table.DataBodyRange(row_index, col_index).Interior.ColorIndex = 6 Then
                issues_to_update(row_index - 1, col_index - 1) = query_update_table.DataBodyRange(row_index, col_index).Value
                issues_to_update(row_index - 1, CHECK_UPDATE_INDEX) = "Update available" 'To avoid looping again, add a variable denoting an update has been made
                Select Case update_type(col_index - 1)
                    Case PUT_POST_NONE
                    Case PUT_GENERIC_FIELD
                        If field_put_generic = False Then
                            field_put_generic = True
                            update_call_count = update_call_count + 1
                            update_calls_for_this_issue = update_calls_for_this_issue + 1
                            issues_to_update(row_index - 1, UPDATE_TYPE_INDEX) = issues_to_update(row_index - 1, UPDATE_TYPE_INDEX) & "field"
                        End If
                    Case PUT_GENERIC_UPDATE
                        If update_put_generic = False Then
                            update_put_generic = True
                            update_call_count = update_call_count + 1
                            update_calls_for_this_issue = update_calls_for_this_issue + 1
                            issues_to_update(row_index - 1, UPDATE_TYPE_INDEX) = issues_to_update(row_index - 1, UPDATE_TYPE_INDEX) & "update"
                        End If
                    Case PUT_GENERIC_UPDATE_REM_ESTIMATE
                        If rem_estimate_put_generic = False Then
                            rem_estimate_put_generic = True
                            update_call_count = update_call_count + 1
                            update_calls_for_this_issue = update_calls_for_this_issue + 1
                            issues_to_update(row_index - 1, UPDATE_TYPE_INDEX) = issues_to_update(row_index - 1, UPDATE_TYPE_INDEX) & "remestimate"
                        End If
                    Case POST_WORKLOG
                        If worklog_post = False Then
                            worklog_post = True
                            update_call_count = update_call_count + 1
                            update_calls_for_this_issue = update_calls_for_this_issue + 1
                            issues_to_update(row_index - 1, UPDATE_TYPE_INDEX) = issues_to_update(row_index - 1, UPDATE_TYPE_INDEX) & "worklog"
                        End If
                    Case POST_STATUS
                        If status_post = False Then
                            status_post = True
                            update_call_count = update_call_count + 1
                            update_calls_for_this_issue = update_calls_for_this_issue + 1
                            issues_to_update(row_index - 1, UPDATE_TYPE_INDEX) = issues_to_update(row_index - 1, UPDATE_TYPE_INDEX) & "status"
                        End If
                    Case Else
                End Select
            End If
        Next
        issues_to_update(row_index - 1, UPDATE_COUNT_FOR_ISSUE) = update_calls_for_this_issue
        'Reset for next issue
        field_put_generic = False
        update_put_generic = False
        rem_estimate_put_generic = False
        worklog_post = False
        status_post = False
    Next
End Sub

Sub WriteResponseAndUpdateProgressBar(row_num As Integer)
    RefreshProgressBar (WorksheetFunction.RoundDown((progress_bar_counter + 1) / (update_call_count + 1) * 100, 0))
    progress_bar_counter = progress_bar_counter + 1
End Sub

Sub UpdateCellColor(issue_row As Integer, issue_field() As String, issue_update_type As String)
    Set query_update_table = ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).ListObjects("JiraQueryUpdateTable")
    Select Case issue_update_type
        Case "field"
            For counter = 0 To UBound(issue_field, 1)
                If update_type(counter) = PUT_GENERIC_FIELD And issue_field(counter) <> "" Then
                    query_update_table.Range(issue_row, counter + 1).Interior.ColorIndex = 4
                End If
            Next
        Case "update"
            For counter = 0 To UBound(issue_field, 1)
                If update_type(counter) = PUT_GENERIC_UPDATE And issue_field(counter) <> "" Then
                    query_update_table.Range(issue_row, counter + 1).Interior.ColorIndex = 4
                End If
            Next
        Case "worklog"
            For counter = 0 To UBound(issue_field, 1)
                If update_type(counter) = POST_WORKLOG And issue_field(counter) <> "" Then
                    query_update_table.Range(issue_row, counter + 1).Interior.ColorIndex = 4
                End If
            Next
        Case "status"
            For counter = 0 To UBound(issue_field, 1)
                If update_type(counter) = POST_STATUS And issue_field(counter) <> "" Then
                    query_update_table.Range(issue_row, counter + 1).Interior.ColorIndex = 4
                End If
            Next
        Case "remestimate"
            For counter = 0 To UBound(issue_field, 1)
                If update_type(counter) = PUT_GENERIC_UPDATE_REM_ESTIMATE And issue_field(counter) <> "" Then
                    query_update_table.Range(issue_row, counter + 1).Interior.ColorIndex = 4
                End If
            Next
        Case Else
            'Do nothing
    End Select
End Sub

'Function to update issues mentioned in table
Sub UpdateModifiedIssues()
    Dim update_start_time As Date
    update_start_time = DateTime.Now 'Timestamp start
    
    Dim issue(QUERY_UPDATE_COLUMN_COUNT - 1) As String 'Zero indexed, hence -1
    progress_bar_counter = 1
    
    GetModifiedIssues
    If update_call_count = 0 Then
        MsgBox "No changes identified in the table", vbOKOnly, "Check issue update to JIRA"
        End
    End If
    
    Dim user_response As Integer
    user_response = MsgBox("This action will update stories specified in the table. Do you confirm?", vbYesNo, "User Confirmation Window")
    Select Case user_response
        'User confirmation
        Case vbYes
            ProgressBar.Show
            For counter = 0 To UBound(issues_to_update, 1)
                issue(ISSUE_KEY_COLUMN) = issues_to_update(counter, ISSUE_KEY_COLUMN) 'ID
                issue(ISSUE_TYPE_COLUMN) = issues_to_update(counter, ISSUE_TYPE_COLUMN) 'Issue type
                issue(ISSUE_STATUS_COLUMN) = issues_to_update(counter, ISSUE_STATUS_COLUMN) 'Status
                issue(ISSUE_SUMMARY_COLUMN) = issues_to_update(counter, ISSUE_SUMMARY_COLUMN) 'Summary
                issue(ISSUE_ASSIGNEE_COLUMN) = issues_to_update(counter, ISSUE_ASSIGNEE_COLUMN) 'Assignee
                issue(ISSUE_EPIC_LINK_COLUMN) = issues_to_update(counter, ISSUE_EPIC_LINK_COLUMN) 'Epic link
                issue(ISSUE_BLOCKED_BY_COLUMN) = issues_to_update(counter, ISSUE_BLOCKED_BY_COLUMN) 'Blocked by
                issue(ISSUE_BLOCKS_COLUMN) = issues_to_update(counter, ISSUE_BLOCKS_COLUMN) 'Blocks
                issue(ISSUE_FIX_VERSION_COLUMN) = issues_to_update(counter, ISSUE_FIX_VERSION_COLUMN) 'Fix version
                issue(ISSUE_PRIORITY_COLUMN) = issues_to_update(counter, ISSUE_PRIORITY_COLUMN) 'Priority
                issue(ISSUE_CUSTOM_FIELD_2_COLUMN) = issues_to_update(counter, ISSUE_CUSTOM_FIELD_2_COLUMN) 'Custom Field 2
                issue(ISSUE_DUE_DATE_COLUMN) = issues_to_update(counter, ISSUE_DUE_DATE_COLUMN) 'Due date
                issue(ISSUE_START_DATE_COLUMN) = issues_to_update(counter, ISSUE_START_DATE_COLUMN) 'Start date
                issue(ISSUE_END_DATE_COLUMN) = issues_to_update(counter, ISSUE_END_DATE_COLUMN) 'End date
                issue(ISSUE_CUSTOM_FIELD_1_COLUMN) = issues_to_update(counter, ISSUE_CUSTOM_FIELD_1_COLUMN) 'Custom Field 1
                issue(ISSUE_COMPONENTS_COLUMN) = issues_to_update(counter, ISSUE_COMPONENTS_COLUMN) 'Component
                issue(ISSUE_ORIGINAL_ESTIMATE_COLUMN) = issues_to_update(counter, ISSUE_ORIGINAL_ESTIMATE_COLUMN) 'Original estimate
                issue(ISSUE_REMAINING_ESTIMATE_COLUMN) = issues_to_update(counter, ISSUE_REMAINING_ESTIMATE_COLUMN) 'Remaining estimate
                issue(ISSUE_TIME_SPENT_COLUMN) = issues_to_update(counter, ISSUE_TIME_SPENT_COLUMN) 'Time spent
                issue(ISSUE_COMMENTS_COLUMN) = issues_to_update(counter, ISSUE_COMMENTS_COLUMN) 'Comment
                
                Dim response_text As Variant
                
                If InStr(issues_to_update(counter, UPDATE_TYPE_INDEX), "field") > 0 And InStr(issues_to_update(counter, 19), "update") > 0 Then
                    Dim UPDATE_ISSUE_JSON_LOCAL As String
                    Call BuildIssueUpdateJson(issue, "field") 'JSON creation for single entry
                    UPDATE_ISSUE_JSON_LOCAL = UPDATE_ISSUE_JSON
                    Call BuildIssueUpdateJson(issue, "update")
                    UPDATE_ISSUE_JSON_LOCAL = UPDATE_ISSUE_JSON_LOCAL & "," & UPDATE_ISSUE_JSON
                    Call SendHttpRequest(API_PUT, issue(ISSUE_KEY_COLUMN), "{" & UPDATE_ISSUE_JSON_LOCAL & "}") 'PUT to jira
                    If InStr(jira_response, "error") = 0 Then
                        Call UpdateCellColor(counter + 2, issue, "field")
                        Call UpdateCellColor(counter + 2, issue, "update")
                    End If
                    WriteResponseAndUpdateProgressBar (counter)
                ElseIf InStr(issues_to_update(counter, UPDATE_TYPE_INDEX), "field") > 0 Then
                    Call BuildIssueUpdateJson(issue, "field") 'JSON creation for single entry
                    Call SendHttpRequest(API_PUT, issue(ISSUE_KEY_COLUMN), "{" & UPDATE_ISSUE_JSON & "}") 'PUT to jira
                    If InStr(jira_response, "error") = 0 Then
                        Call UpdateCellColor(counter + 2, issue, "field")
                    End If
                    WriteResponseAndUpdateProgressBar (counter)
                ElseIf InStr(issues_to_update(counter, UPDATE_TYPE_INDEX), "update") > 0 Then
                    Call BuildIssueUpdateJson(issue, "update") 'JSON creation for single entry
                    Call SendHttpRequest(API_PUT, issue(ISSUE_KEY_COLUMN), "{" & UPDATE_ISSUE_JSON & "}") 'PUT to jira
                    If InStr(jira_response, "error") = 0 Then
                        Call UpdateCellColor(counter + 2, issue, "update")
                    End If
                    WriteResponseAndUpdateProgressBar (counter)
                End If
                If InStr(issues_to_update(counter, UPDATE_TYPE_INDEX), "remestimate") > 0 Then
                    Call BuildIssueUpdateJson(issue, "estimate") 'JSON creation for single entry
                    Call SendHttpRequest(API_PUT, issue(ISSUE_KEY_COLUMN), UPDATE_ISSUE_JSON) 'PUT to jira
                    If InStr(jira_response, "error") = 0 Then
                        Call UpdateCellColor(counter + 2, issue, "remestimate")
                    End If
                    WriteResponseAndUpdateProgressBar (counter)
                End If
                If InStr(issues_to_update(counter, UPDATE_TYPE_INDEX), "worklog") > 0 Then
                    Call BuildIssueUpdateJson(issue, "worklog") 'JSON creation for single entry
                    Call SendHttpRequest(API_POST, issue(ISSUE_KEY_COLUMN) & "/worklog", UPDATE_ISSUE_JSON) 'POST to jira
                    If InStr(jira_response, "error") = 0 Then
                        Call UpdateCellColor(counter + 2, issue, "worklog")
                    End If
                    WriteResponseAndUpdateProgressBar (counter)
                End If
                If InStr(issues_to_update(counter, UPDATE_TYPE_INDEX), "status") > 0 Then
                    Call BuildIssueUpdateJson(issue, "status") 'JSON creation for single entry
                    Call SendHttpRequest(API_POST, issue(ISSUE_KEY_COLUMN) & "/transitions?expand=transitions.fields", UPDATE_ISSUE_JSON) 'POST to jira
                    If InStr(jira_response, "error") = 0 Then
                        Call UpdateCellColor(counter + 2, issue, "status")
                    End If
                    WriteResponseAndUpdateProgressBar (counter)
                End If
            Next
            Unload ProgressBar
    End Select
    ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range("B6") = DateDiff("s", update_start_time, DateTime.Now) & " sec" 'Write time to create
    EnableTracking
    'Bulk update logic
        'Dim multi_issue_string as String
        'multi_issue_string = ""
        'multi_issue_string = "{""issueUpdates"": ["
        'multi_issue_string = multi_issue_string & CREATE_ISSUE_JSON
        'If counter < ISSUE_COUNT - 1 Then
        '   multi_issue_string = multi_issue_string & ","
        'End If
        'multi_issue_string = multi_issue_string & "]}"
        'Call SendHttpRequest(API_POST, "bulk", multi_issue_string)
End Sub
