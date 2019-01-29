Attribute VB_Name = "CreateIssue"
'Used across functions in CreateIssue module
Dim issues_to_create() As String

'Get details of owners provided in the sheet
Sub GetOwners()
    owner_1 = ThisWorkbook.Worksheets(SHEET_CREATE).Range("B6").Value
    owner_2 = ThisWorkbook.Worksheets(SHEET_CREATE).Range("B7").Value
    owner_3 = ThisWorkbook.Worksheets(SHEET_CREATE).Range("B8").Value
    owner_4 = ThisWorkbook.Worksheets(SHEET_CREATE).Range("B9").Value
    owner_5 = ThisWorkbook.Worksheets(SHEET_CREATE).Range("B10").Value
End Sub

'This is a template for display. Each epic is expected to have 22 stories linked by default
Sub PopulateTemplateToTable()
    GetOwners
    PopulateIssueTemplate
    
    Dim LABELS As String
    Dim CUSTOM_FIELD As String
    Dim FIX_VERSION As String
    
    FIX_VERSION = ThisWorkbook.Worksheets(SHEET_CREATE).Range("B3").Value
    LABELS = ThisWorkbook.Worksheets(SHEET_CREATE).Range("B4").Value
    CUSTOM_FIELD = ThisWorkbook.Worksheets(SHEET_CREATE).Range("B5").Value
    
    For row = 0 To issue_count - 1
        template(row, 0) = "Story"
        template(row, 4) = FIX_VERSION
        template(row, 8) = LABELS
        template(row, 9) = CUSTOM_FIELD
    Next
    
    Set issue_table = ThisWorkbook.Worksheets(SHEET_CREATE).ListObjects(1) 'JiraCreateIssueTable
    On Error Resume Next
    issue_table.DataBodyRange.Delete
    ThisWorkbook.Worksheets(SHEET_CREATE).Range("B14:L" & (14 + issue_count - 1)).Value = template 'Zero indexed, hence -1
End Sub

'Get details of issues to be generated from table and store in program context
Sub GetIssueDetailsFromTable()
    Dim issue_table As ListObject
    Set issue_table = ThisWorkbook.Worksheets(SHEET_CREATE).ListObjects(1)
    Dim table_data As Variant
    table_data = issue_table.DataBodyRange.Rows
    
    ReDim issues_to_create(issue_table.Range.Rows.Count - 2, issue_table.Range.Columns.Count - 2) As String
    Dim epic_names() As String
    ReDim epic_names(issue_table.Range.Rows.Count - 2, 1) As String
    
    'Logic to optimize JIRA gets
    Dim duplicate_entry As Boolean
    duplicate_entry = False
    epic_names(0, 0) = table_data(1, 7) 'Get epic key of first entry
    If epic_names(0, 0) <> "" Then
        epic_names(0, 1) = GetEpicName(CStr(table_data(1, 7)))
    End If
    
    For row_index = LBound(table_data) To UBound(table_data)
        epic_names(row_index - 1, 0) = table_data(row_index, 7)
        'Re-use epic name
        If epic_names(row_index - 1, 1) = "" Then
            For inner_loop = LBound(table_data) To row_index - 1
                If epic_names(row_index - 1, 0) = epic_names(inner_loop - 1, 0) Then
                    epic_names(row_index - 1, 1) = epic_names(inner_loop - 1, 1)
                    duplicate_entry = True
                    Exit For
                End If
                If inner_loop = row_index - 1 Then
                    duplicate_entry = False
                End If
            Next
            'Get epic name from JIRA
            If duplicate_entry = False And epic_names(row_index - 1, 0) <> "" Then
                epic_names(row_index - 1, 1) = GetEpicName(epic_names(row_index - 1, 0))
                duplicate_entry = True
            End If
        End If
    Next
    
    'Convert newline [Chr(10)] to "\n"
    For row_index = LBound(table_data) To UBound(table_data)
        For col_index = LBound(table_data, 2) + 1 To UBound(table_data, 2)
            issues_to_create(row_index - 1, col_index - 2) = Replace(table_data(row_index, col_index), Chr(10), "\n")
        Next col_index
    Next row_index
    
    'Check user prefix
    Dim user_prefix As String
    user_prefix = ThisWorkbook.Worksheets(SHEET_CREATE).Range("B12").Value
    If user_prefix <> "" Then
        user_prefix = user_prefix & ": "
    End If
    'Update summary with epic name and user provided prefix in Cell B12
    For row_index = LBound(table_data) To UBound(table_data)
        If epic_names(row_index - 1, 1) <> "" Then
            issues_to_create(row_index - 1, 1) = epic_names(row_index - 1, 1) & ": " & user_prefix & table_data(row_index, 3)
        Else: issues_to_create(row_index - 1, 1) = user_prefix & table_data(row_index, 3)
        End If
    Next row_index
End Sub

'Generate JIRA post parameter of multi-select fields
Function GetMultiValueString(field_type As String, values As String)
    Dim multi_value_string As String
    multi_value_string = ""
    Dim trimmed_values As String
    trimmed_values = Replace(values, " ", "")
    Dim value_array() As String
    value_array = Split(values, ",")
    
    For counter = 0 To UBound(value_array)
        multi_value_string = multi_value_string & "{""" & field_type & """:""" & WorksheetFunction.Trim(value_array(counter)) & """}"
        If counter < UBound(value_array) Then
            multi_value_string = multi_value_string & ","
        End If
    Next
    GetMultiValueString = multi_value_string
End Function

'Creates JSON to be sent to JIRA
Function BuildStoryCreationJson(issue_field() As String) As String
    Dim ISSUE_HEADER As String
    Dim PROJECT As String
    Dim ISSUE_TYPE As String
    Dim SUMMARY As String
    Dim ASSIGNEE As String
    Dim CUSTOM_FIELD_0 As String
    Dim CUSTOM_FIELD_1 As String
    Dim CUSTOM_FIELD_2 As String
    Dim FIX_VERSION As String
    Dim PRIORITY As String
    Dim DUE_DATE As String
    Dim description As String
    Dim LABELS As String
    Dim ISSUE_VIEWERS As String
    Dim ISSUE_FOOTER As String
    Dim separator As String
    separator = ","

    ISSUE_HEADER = "{""fields"":{"
    PROJECT = """project"":{""key"":""" & ThisWorkbook.Worksheets(SHEET_CREATE).Range("B1").Value & """}"
    ISSUE_TYPE = """issuetype"": {""name"":""" & issue_field(0) & """}"
    SUMMARY = """summary"":""" & issue_field(1) & """"
    ASSIGNEE = """assignee"":{""name"":""" & issue_field(2) & """}"
    
    'Add CUSTOM_FIELD_0 parameter only if provided
    Dim CUSTOM_FIELD_0_values As String
    CUSTOM_FIELD_0_values = GetMultiValueString("value", issue_field(3))
    If CUSTOM_FIELD_0_values <> "" Then
        CUSTOM_FIELD_0 = """customfield_xxxxx"": [" & CUSTOM_FIELD_0_values & "]"
    Else: CUSTOM_FIELD_0 = ""
    End If
    
    'Add fix version parameter only if provided
    Dim fix_version_values As String
    fix_version_values = GetMultiValueString("name", issue_field(4))
    If fix_version_values <> "" Then
        FIX_VERSION = """fixVersions"": [" & fix_version_values & "]"
    Else: FIX_VERSION = ""
    End If
    
    'Add epic link parameter only if provided
    If issue_field(5) <> "" Then
        CUSTOM_FIELD_2 = """customfield_xxxxx"":""" & issue_field(5) & """"
    Else: CUSTOM_FIELD_2 = ""
    End If
    PRIORITY = """priority"": {""name"":""" & issue_field(6) & """}"
    description = """description"":""" & issue_field(7) & """"
    LABELS = """labels"": [""" & Replace(WorksheetFunction.Trim(issue_field(8)), ",", """,""") & """]"
    
    'Add CUSTOM_FIELD_1 parameter only if provided
    Dim CUSTOM_FIELD_1_values As String
    CUSTOM_FIELD_1_values = GetMultiValueString("value", issue_field(9))
    If CUSTOM_FIELD_1_values <> "" Then
        CUSTOM_FIELD_1 = """customfield_xxxxx"": [" & CUSTOM_FIELD_1_values & "]"
    Else: CUSTOM_FIELD_1 = ""
    End If
    
    'Add due date parameter only if provided
    If issue_field(10) <> "" Then
        DUE_DATE = """duedate"":""" & Format(issue_field(10), "yyyy-mm-dd") & """"
    End If
    ISSUE_FOOTER = "}}"

    CREATE_ISSUE_JSON = ISSUE_HEADER & _
                PROJECT & separator & _
                ISSUE_TYPE & separator & _
                SUMMARY & separator & _
                ASSIGNEE & separator & _
                PRIORITY & separator & _
                description & separator & _
                LABELS
                
    'Add parameters only if available. Empty parameter value will lead to JIRA API call failure
    If CUSTOM_FIELD_2 <> "" Then
        CREATE_ISSUE_JSON = CREATE_ISSUE_JSON & separator & CUSTOM_FIELD_2
    End If
    If CUSTOM_FIELD_0 <> "" Then
        CREATE_ISSUE_JSON = CREATE_ISSUE_JSON & separator & CUSTOM_FIELD_0
    End If
    If FIX_VERSION <> "" Then
        CREATE_ISSUE_JSON = CREATE_ISSUE_JSON & separator & FIX_VERSION
    End If
    If CUSTOM_FIELD_1 <> "" Then
        CREATE_ISSUE_JSON = CREATE_ISSUE_JSON & separator & CUSTOM_FIELD_1
    End If
    If DUE_DATE <> "" Then
        CREATE_ISSUE_JSON = CREATE_ISSUE_JSON & separator & DUE_DATE
    End If
    
    CREATE_ISSUE_JSON = CREATE_ISSUE_JSON & ISSUE_FOOTER
    
    'Return the constructed JSON
    BuildIssueCreationString = CREATE_ISSUE_JSON
End Function

'Function to create issues mentioned in table
Sub CreateJiraIssues()
    start_time = DateTime.Now 'Timestamp start
    Dim issue(CREATE_STORY_FIELD_COUNT - 1) As String 'Zero indexed, hence -1
    
    Dim user_response As Integer
    user_response = MsgBox("This action will create stories specified in the table. Do you confirm?", vbYesNo, "User confirmation for issue creation")
    Select Case user_response
        'User confirmation
        Case vbYes
            GetIssueDetailsFromTable
            ProgressBar.Show
            For counter = 0 To UBound(issues_to_create, 1)
                issue(0) = issues_to_create(counter, 0)
                issue(1) = issues_to_create(counter, 1)
                issue(2) = issues_to_create(counter, 2)
                issue(3) = issues_to_create(counter, 3)
                issue(4) = issues_to_create(counter, 4)
                issue(5) = issues_to_create(counter, 5)
                issue(6) = issues_to_create(counter, 6)
                issue(7) = issues_to_create(counter, 7)
                issue(8) = issues_to_create(counter, 8)
                issue(9) = issues_to_create(counter, 9)
                issue(10) = issues_to_create(counter, 10)
                Dim row_entry As Range
                Set row_entry = ThisWorkbook.Worksheets(SHEET_CREATE).Range("A" & counter + 14 & ":L" & counter + 14)
                If WorksheetFunction.CountA(row_entry.EntireRow) <> 0 Then 'Error handling to ensure we don't call JIRA create for empty items
                    Call BuildStoryCreationJson(issue) 'JSON creation for single entry
                    Call SendHttpRequest(API_POST, "", CREATE_ISSUE_JSON) 'POST to jira
                    Dim response_text As Variant
                    On Error GoTo issue_creation_handler 'Throws error if CallByName fails
                    If InStr(jira_response, "error") = 0 And jira_response <> "" Then
                       Set response_text = scriptControl.Eval("(" + jira_response + ")")
                       jira_response = CallByName(response_text, "key", VbGet) 'Created issue key
                    End If
                End If
issue_creation_handler:
                ThisWorkbook.Worksheets(SHEET_CREATE).Range("A" & counter + 14).Value = jira_response 'Write to table
                RefreshProgressBar (WorksheetFunction.RoundDown((counter + 1) / (UBound(issues_to_create, 1) + 1) * 100, 0))
            Next
            Unload ProgressBar
    End Select
    ThisWorkbook.Worksheets(SHEET_CREATE).Range("B11") = DateDiff("s", start_time, DateTime.Now) & " sec" 'Capture time to create
    
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
