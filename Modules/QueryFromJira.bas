Attribute VB_Name = "QueryFromJira"

'Query list of all Proj1 users considering Proj1 is accessible across all
Sub GetUsers()
    Dim user_table As ListObject
    Set user_table = ThisWorkbook.Worksheets(SHEET_IDS).ListObjects("Proj1Users")
    
    'Clear table
    If Not user_table.DataBodyRange Is Nothing Then
        user_table.AutoFilter.ShowAllData
        user_table.DataBodyRange.Delete
    End If
    
    Dim user_ids() As Variant
    Dim user_display_names() As Variant
    Dim total_users As Integer
    total_users = 0
    'Max limit of 10000 users
    For counter = 0 To 9
        GetHttpRequest (JIRA_API_BASE_URL & "user/assignable/search?project=Proj1&maxResults=10000&startAt=" & (counter * 1000))
        If jira_json = "" Then
            Exit For
        End If
        total_users = total_users + CallByName(jira_json, "length", VbGet)
        ReDim Preserve user_ids(total_users - 1)
        ReDim Preserve user_display_names(total_users - 1)
        Dim index_0 As String
        Dim index_1 As String
        index_0 = "name"
        index_1 = "displayName"
        For Count = (counter * 1000) To UBound(user_ids)
            user_ids(Count) = CallByName(CallByName(jira_json, Count - (counter * 1000), VbGet), index_0, VbGet)
            user_display_names(Count) = CallByName(CallByName(jira_json, Count - (counter * 1000), VbGet), index_1, VbGet)
        Next
    Next counter
    ThisWorkbook.Worksheets(SHEET_IDS).Range("AM2:AM" & (total_users + 1)).Value = Application.Transpose(user_display_names)
    ThisWorkbook.Worksheets(SHEET_IDS).Range("AN2:AN" & (total_users + 1)).Value = Application.Transpose(user_ids)
    user_table.Resize user_table.Range.Resize(total_users + 1)
End Sub

'Get IDs of user queried fields
Sub GetIds(get_type As String, start_column As String, end_column As String)
    GetHttpRequest (JIRA_API_BASE_URL & get_type)
    Dim queried_ids() As Variant
    ReDim queried_ids(CallByName(jira_json, "length", VbGet) - 1, 1)
    Dim index_0 As String
    Dim index_1 As String
    index_0 = "name"
    index_1 = "id"
    If get_type = "project" Then
        index_0 = "key"
    End If
    For Count = 0 To UBound(queried_ids)
        queried_ids(Count, 0) = CallByName(CallByName(jira_json, Count, VbGet), index_0, VbGet)
        queried_ids(Count, 1) = CallByName(CallByName(jira_json, Count, VbGet), index_1, VbGet)
    Next
    ThisWorkbook.Worksheets(SHEET_IDS).Range(start_column & "2:" & end_column & (UBound(queried_ids) + 2)).Value = queried_ids()
End Sub

'Wrapper to get few fields and IDs
Sub GetIdsFromJira()
    Call GetIds("issuetype", "A", "B")
    Call GetIds("project", "D", "E")
    Call GetIds("status", "G", "H")
    Call GetIds("project/Proj3/versions", "O", "P")
    Call GetIds("field", "R", "S")
    Call GetIds("filter/favourite", "U", "V")
    Call GetIds("resolution", "AS", "AT")
End Sub

'Function to get allowed values of user specified field
Sub GetAllowedFieldValues(field As String, query_type As String, start_column As String, end_column As String)
    GetHttpRequest (JIRA_API_ISSUE_URL & "createmeta?projectKeys=Proj2&expand=projects.issuetypes.fields")
    Set jira_json = CallByName(CallByName(CallByName(CallByName(CallByName(jira_json, "projects", VbGet), 0, VbGet), "issuetypes", VbGet), 1, VbGet), "fields", VbGet)
    Set jira_json = CallByName(CallByName(jira_json, field, VbGet), "allowedValues", VbGet)
    Dim field_length As Integer
    field_length = Len(jira_json)
    field_length = field_length - 15
    field_length = field_length / 16
    Dim queried_ids() As Variant
    ReDim queried_ids(field_length, 1) As Variant
    For Count = 0 To field_length
        queried_ids(Count, 0) = CallByName(CallByName(jira_json, Count, VbGet), query_type, VbGet)
        queried_ids(Count, 1) = CallByName(CallByName(jira_json, Count, VbGet), "id", VbGet)
    Next
    ThisWorkbook.Worksheets(SHEET_IDS).Range(start_column & "2:" & end_column & (UBound(queried_ids) + 2)).Value = queried_ids()
End Sub

'Wrapper to get allowed values of few fields
Sub GetAllowedValues()
    Call GetAllowedFieldValues(CUSTOM_FIELD_2, "value", "AJ", "AK")
    Call GetAllowedFieldValues(CUSTOM_FIELD_1, "value", "AG", "AH")
    Call GetAllowedFieldValues("components", "name", "AD", "AE")
    Call GetAllowedFieldValues("priority", "name", "AA", "AB")
    Call GetProj2FixVersion
    Call GetAllowedFieldValues(FLAG_FIELD, "value", "X", "Y")
End Sub

'Function to get fix version values of user specified field
Sub GetProj2FixVersion()
    GetHttpRequest (JIRA_API_ISSUE_URL & "createmeta?projectKeys=Proj2&expand=projects.issuetypes.fields")
    Set jira_json = CallByName(CallByName(CallByName(CallByName(CallByName(jira_json, "projects", VbGet), 0, VbGet), "issuetypes", VbGet), 1, VbGet), "fields", VbGet)
    Set jira_json = CallByName(CallByName(jira_json, "fixVersions", VbGet), "allowedValues", VbGet)
    Dim field_length As Integer
    field_length = Len(jira_json)
    field_length = field_length - 15
    field_length = field_length / 16
    Dim queried_ids() As Variant
    ReDim queried_ids(field_length, 3) As Variant
    For Count = 0 To field_length
        On Error Resume Next
        queried_ids(Count, 0) = CallByName(CallByName(jira_json, Count, VbGet), "name", VbGet)
        queried_ids(Count, 1) = CallByName(CallByName(jira_json, Count, VbGet), "id", VbGet)
        queried_ids(Count, 2) = CallByName(CallByName(jira_json, Count, VbGet), "startDate", VbGet)
        queried_ids(Count, 3) = CallByName(CallByName(jira_json, Count, VbGet), "releaseDate", VbGet)
    Next
    ThisWorkbook.Worksheets(SHEET_IDS).Range("J2:M" & (UBound(queried_ids) + 2)).Value = queried_ids()
End Sub

'Refresh IDs sheet
Sub RefreshIds()
    Dim table_count As Integer
    table_count = ThisWorkbook.Worksheets(SHEET_IDS).ListObjects.Count
    For counter = 1 To table_count
        If Not ThisWorkbook.Worksheets(SHEET_IDS).ListObjects(counter).DataBodyRange Is Nothing Then
            ThisWorkbook.Worksheets(SHEET_IDS).ListObjects(counter).AutoFilter.ShowAllData
            ThisWorkbook.Worksheets(SHEET_IDS).ListObjects(counter).DataBodyRange.Delete
        End If
    Next counter
    GetIdsFromJira
    GetAllowedValues
    GetUsers
End Sub

'Get description of the JIRA key in column A
Sub GetDescription()
    Dim jira_key As String
    jira_key = ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range("A" & ActiveCell.row).Value
    
    'Return if no JIRA key present
    If jira_key <> "" Then
        GetHttpRequest (JIRA_API_ISSUE_URL & jira_key & "?fields=description")
        
        'Show default message if no description
        If FieldHasValue(jira_json, "fields/description") Then
            JiraDescription.DescriptionText.Value = CallByName(CallByName(jira_json, "fields", VbGet), "description", VbGet)
        Else: JiraDescription.DescriptionText.Value = NO_DESCRIPTION_STRING
        End If
        
        'Show description userform
        JiraDescription.Caption = "Description for " & jira_key
        If JiraDescription.Visible = False Then
            JiraDescription.Show
        End If
        
    Else: MsgBox "No JIRA key available. Please ensure you have the correct row selected in the table", vbOKOnly, "Invalid JIRA key"
    End If
End Sub

'Get acceptance criteria of the JIRA key in column A
Sub GetAcceptanceCriteria()
    Dim jira_key As String
    jira_key = ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range("A" & ActiveCell.row).Value
    
    'Return if no JIRA key present
    If jira_key <> "" Then
        GetHttpRequest (JIRA_API_ISSUE_URL & jira_key & "?fields=customfield_11609")
        
        'Show default message if no description
        If FieldHasValue(jira_json, "fields/customfield_xxxxx") Then
            AcceptanceCriteria.AcceptanceCriteriaText.Value = CallByName(CallByName(jira_json, "fields", VbGet), "customfield_xxxxx", VbGet)
        Else: AcceptanceCriteria.AcceptanceCriteriaText.Value = NO_ACCEPTANCE_CRITERIA_STRING
        End If
        
        'Show description userform
        AcceptanceCriteria.Caption = "Acceptance Criteria for " & jira_key
        If AcceptanceCriteria.Visible = False Then
            AcceptanceCriteria.Show
        End If
        
    Else: MsgBox "No JIRA key available. Please ensure you have the correct row selected in the table", vbOKOnly, "Invalid JIRA key"
    End If
End Sub

'Get worklog of the JIRA key in column A
Sub GetWorklog()
    Dim jira_key As String
    jira_key = ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range("A" & ActiveCell.row).Value
    
    'Return if no JIRA key present
    If jira_key <> "" Then
        GetHttpRequest (JIRA_API_ISSUE_URL & jira_key & "/worklog")
        
        Dim worklogs_from_jira As String
        worklogs_from_jira = ""
        
        Dim all_worklogs As Object
        Set all_worklogs = CallByName(jira_json, "worklogs", VbGet)
        
        Dim current_worklog As Object
        Dim created_date_time As String
        
        Dim Worklog_count As Integer
        Worklog_count = CallByName(jira_json, "total", VbGet)
        
        For counter = 0 To Worklog_count - 1
            'Extract one by one
            Set current_worklog = CallByName(all_worklogs, counter, VbGet)
            'Get autohr name
            worklogs_from_jira = worklogs_from_jira & CallByName(CallByName(current_worklog, "author", VbGet), "displayName", VbGet) & ": "
            'Get worklog date
            created_date_time = CallByName(current_worklog, "created", VbGet)
            'Extract date and time from worklog date
            Worklog.WorklogIDs.AddItem CallByName(current_worklog, "id", VbGet)
            worklogs_from_jira = worklogs_from_jira & "Log Time: " & Left(created_date_time, 10) & " " & Left(Right(created_date_time, 17), 8)
            worklogs_from_jira = worklogs_from_jira & vbTab & vbTab & vbTab & "ID: " & CallByName(current_worklog, "id", VbGet) & vbCrLf
            worklogs_from_jira = worklogs_from_jira & ".........................." & vbCrLf
            If FieldHasValue(current_worklog, "comment") Then
                worklogs_from_jira = worklogs_from_jira & "Time spent: " & CallByName(current_worklog, "timeSpent", VbGet) & vbCrLf & " Comment: " & CallByName(current_worklog, "comment", VbGet) & vbCrLf
            Else: worklogs_from_jira = worklogs_from_jira & "Time spent: " & CallByName(current_worklog, "timeSpent", VbGet) & vbCrLf & " Comment: None" & vbCrLf
            End If
            worklogs_from_jira = worklogs_from_jira & "===================================================================="
            If counter < Worklog_count - 1 Then
                'New lines except for last worklogs
                worklogs_from_jira = worklogs_from_jira & vbCrLf & vbCrLf
            End If
        Next
        
        'Show default message if no worklog
        If CallByName(jira_json, "total", VbGet) > 0 Then
            Worklog.WorklogText.Text = worklogs_from_jira
        Else: Worklog.WorklogText.Text = "No worklogs available"
        End If
        
        'Show userform
        Worklog.Caption = "Worklogs for " & jira_key
        If Worklog.Visible = False Then
            Worklog.Show
        End If
    Else: MsgBox "No JIRA key available. Please ensure you have the correct row selected in the table", vbOKOnly, "Invalid JIRA key"
    End If
End Sub

'Get comment of the JIRA key in column A
Sub GetComment()
    Dim jira_key As String
    jira_key = ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range("A" & ActiveCell.row).Value
    
    'Return if no JIRA key present
    If jira_key <> "" Then
        GetHttpRequest (JIRA_API_ISSUE_URL & jira_key & "/comment")
        
        Dim comments_from_jira As String
        comments_from_jira = ""
        
        Dim all_comments As Object
        Set all_comments = CallByName(jira_json, "comments", VbGet)
        
        Dim current_comment As Object
        Dim created_date_time As String
        
        Dim comment_count As Integer
        comment_count = CallByName(jira_json, "total", VbGet)
        
        For counter = 0 To comment_count - 1
            'Extract one by one
            Set current_comment = CallByName(all_comments, counter, VbGet)
            'Get commenter name
            comments_from_jira = comments_from_jira & CallByName(CallByName(current_comment, "author", VbGet), "displayName", VbGet) & ": "
            'Get comment date
            created_date_time = CallByName(current_comment, "created", VbGet)
            'Extract date and time from comment date
            Comments.CommentIDs.AddItem CallByName(current_comment, "id", VbGet)
            comments_from_jira = comments_from_jira & "Time: " & Left(created_date_time, 10) & " " & Left(Right(created_date_time, 17), 8)
            comments_from_jira = comments_from_jira & vbTab & vbTab & vbTab & "ID: " & CallByName(current_comment, "id", VbGet) & vbCrLf
            comments_from_jira = comments_from_jira & ".........................." & vbCrLf
            comments_from_jira = comments_from_jira & CallByName(current_comment, "body", VbGet) & vbCrLf
            comments_from_jira = comments_from_jira & "===================================================================="
            If counter < comment_count - 1 Then
                'New lines except for last comment
                comments_from_jira = comments_from_jira & vbCrLf & vbCrLf
            End If
        Next
        
        'Show default message if no comments
        If CallByName(jira_json, "total", VbGet) > 0 Then
            Comments.CommentText.Text = comments_from_jira
        Else: Comments.CommentText.Text = "No comments available"
        End If
        
        'Show in userform
        Comments.Caption = "Comments for " & jira_key
        If Comments.Visible = False Then
            Comments.Show
        End If
    Else: MsgBox "No JIRA key available. Please ensure you have the correct row selected in the table", vbOKOnly, "Invalid JIRA key"
    End If
End Sub

Sub GetCreateMeta()
    If SESSION_ID = "" Then
        SESSION_ID = GetSessionId()
    End If
    
    Set scriptControl = CreateObject("MSScriptControl.ScriptControl")
    scriptControl.Language = "JScript"
    
    With JiraService
    .Open "GET", "https://<jira-url>/rest/api/latest/issue/createmeta?projectKeys=Proj2&issuetypeNames=Type&expand=projects.issuetypes.fields", False
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "Accept", "application/json"
        .setRequestHeader "X-Atlassian-Token", "nocheck"
        .setRequestHeader "Set-Cookie", SESSION_ID
        .send
        If .status = 401 Then 'Handle access denied error
            ThisWorkbook.Worksheets(WORKSHEET_NAME).Range("A1").Value = ""
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
            Set jira_json = CallByName(jira_json, "projects", VbGet)
            Set jira_json = CallByName(jira_json, 0, VbGet)
            Set jira_json = CallByName(jira_json, "issuetypes", VbGet)
            Set jira_json = CallByName(jira_json, 0, VbGet)
            Set jira_json = CallByName(jira_json, "fields", VbGet)
        End If
        .abort
    End With
    Dim queried_ids() As Variant
    ReDim queried_ids(CallByName(jira_json, "length", VbGet) - 1, 1)
    Dim field_count As Integer
    field_count = Len(jira_json)
    Dim index_0 As String
    Dim index_1 As String
    index_0 = "name"
    index_1 = "id"
    If get_type = "project" Then
        index_0 = "key"
    End If
    For Count = 0 To UBound(queried_ids)
        queried_ids(Count, 0) = CallByName(CallByName(jira_json, Count, VbGet), index_0, VbGet)
        queried_ids(Count, 1) = CallByName(CallByName(jira_json, Count, VbGet), index_1, VbGet)
    Next
    ThisWorkbook.Worksheets(SHEET_IDS).Range(start_column & "2:" & end_column & (UBound(queried_ids) + 2)).Value = queried_ids()
End Sub
