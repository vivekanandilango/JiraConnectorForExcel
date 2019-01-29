Attribute VB_Name = "Globals"
Public jira_response As Variant
Public CREATE_ISSUE_JSON As String
Public UPDATE_ISSUE_JSON As String
Public owner_1 As String
Public owner_2 As String
Public owner_5 As String
Public owner_3 As String
Public owner_4 As String
Public template(CREATE_ISSUE_COUNT, CREATE_STORY_FIELD_COUNT) As String

Public JiraService As New MSXML2.XMLHTTP60
Public scriptControl As Object
Public SESSION_ID As String
Public jira_json As Object
Public jira_issue_json As Object
Public jira_json_fields As Object
Public jira_entry() As Variant
Public table_range As Range
Public query_update_table As ListObject
Public create_table As ListObject
Public response_text As String
Public start_time As Date
Public query_row_count As Integer
Public query_col_count As Integer
Public track_changes As Boolean
Public custom_field_2_count As Integer
Public custom_field_1_count As Integer
Public fix_version_count As Integer
Public issue_links_count As Integer
Public holidays As Variant

Public return_value As Integer
Public jira_key As String

Public update_type(QUERY_UPDATE_COLUMN_COUNT) As Integer
Public update_call_count As Long

Public dict_open_states As New Dictionary
