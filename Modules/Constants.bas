Public Const CREATE_ISSUE_COUNT As Integer = 22
Public Const CREATE_STORY_FIELD_COUNT As Integer = 11
Public Const SHEET_CREATE As String = "Create"
Public Const SHEET_IDS As String = "IDs"
Public Const SHEET_QUERY_UPDATE As String = "Query_Update"
Public Const SHEET_HOLIDAYS As String = "Holidays"

Public Const WORK_HOURS_PER_DAY As Integer = 8
Public Const QUERY_UPDATE_COLUMN_COUNT As Integer = 24
Public Const QUERY_UPDATE_HEADER_RANGE As String = "A8:X8"
Public Const QUERY_UPDATE_DATABODYRANGE_COL As String = "A9:X"
Public Const QUERY_UPDATE_WRITE_COL As String = "A9:X"
Public Const QUERY_UPDATE_ROW_OFFSET As Integer = 8
Public Const CELL_QUERY_UPDATE_DATABODYRANGE_START As String = "A9"
Public Const CELL_ISSUE_TYPE_DROPDOWN As String = "B9"
Public Const CELL_STATUS_DROPDOWN As String = "C9"
Public Const CELL_USER_ID_DROPDOWN As String = "E9"
Public Const CELL_FIX_VERSION_DROPDOWN As String = "I9"
Public Const CELL_PRIORITY_DROPDOWN As String = "J9"
Public Const CELL_CUSTOM_FIELD_1_DROPDOWN As String = "O9"
Public Const CELL_CUSTOM_FIELD_2_DROPDOWN As String = "K9"
Public Const CELL_CUSTOM_FIELD_3_DROPDOWN As String = "P9"
Public Const CELL_TREND_DROPDOWN As String = "V9"
Public Const CELL_ASSESSMENT_DROPDOWN As String = "W9"
Public Const JIRA_RESPONSE_COLUMN As String = "V"

Public Const API_PUT As String = "PUT"
Public Const API_POST As String = "POST"
Public Const JIRA_API_BASE_URL As String = "https://<jira-url>/rest/api/latest/"
Public Const JIRA_API_ISSUE_URL As String = JIRA_API_BASE_URL & "issue/"
Public Const JIRA_API_SEARCH_URL As String = JIRA_API_BASE_URL & "search?jql="
Public Const JIRA_API_FILTER_URL As String = JIRA_API_BASE_URL & "filter/"
Public Const JIRA_API_AUTH_URL As String = "https://<jira-url>/rest/auth/1/session"

Public Const CUSTOM_FIELD_0 As String = "customfield_xxxxx"
Public Const CUSTOM_FIELD_1 As String = "customfield_xxxxx"
Public Const CUSTOM_FIELD_2 As String = "customfield_xxxxx"
Public Const CUSTOM_FIELD_3 As String = "customfield_xxxxx"
Public Const CUSTOM_FIELD_4 As String = "customfield_xxxxx"
Public Const CUSTOM_FIELD_5 As String = "customfield_xxxxx"
Public Const CUSTOM_FIELD_6 As String = "customfield_xxxxx"
Public Const CUSTOM_FIELD_7 As String = "customfield_xxxxx"
Public Const CUSTOM_FIELD_8 As String = "customfield_xxxxx"

'Field column indices (zero indexed)
Public Const ISSUE_KEY_COLUMN As Integer = 0
Public Const ISSUE_TYPE_COLUMN As Integer = 1
Public Const ISSUE_STATUS_COLUMN As Integer = 2
Public Const ISSUE_SUMMARY_COLUMN As Integer = 3
Public Const ISSUE_ASSIGNEE_COLUMN As Integer = 4
Public Const ISSUE_EPIC_LINK_COLUMN As Integer = 5
Public Const ISSUE_BLOCKED_BY_COLUMN As Integer = 6
Public Const ISSUE_BLOCKS_COLUMN As Integer = 7
Public Const ISSUE_FIX_VERSION_COLUMN As Integer = 8
Public Const ISSUE_PRIORITY_COLUMN As Integer = 9
Public Const ISSUE_CUSTOM_FIELD_2_COLUMN As Integer = 10
Public Const ISSUE_DUE_DATE_COLUMN As Integer = 11
Public Const ISSUE_START_DATE_COLUMN As Integer = 12
Public Const ISSUE_END_DATE_COLUMN As Integer = 13
Public Const ISSUE_CUSTOM_FIELD_1_COLUMN As Integer = 14
Public Const ISSUE_COMPONENTS_COLUMN As Integer = 15
Public Const ISSUE_ORIGINAL_ESTIMATE_COLUMN As Integer = 16
Public Const ISSUE_REMAINING_ESTIMATE_COLUMN As Integer = 17
Public Const ISSUE_TIME_SPENT_COLUMN As Integer = 18
Public Const ISSUE_COMMENTS_COLUMN As Integer = 19
Public Const ISSUE_BANDWIDTH_COLUMN As Integer = 20
Public Const ISSUE_TREND_COLUMN As Integer = 21
Public Const ISSUE_ASSESSMENT_COLUMN As Integer = 22
Public Const ISSUE_BEST_END_COLUMN As Integer = 23
Public Const UPDATE_TYPE_INDEX As Integer = 24
Public Const CHECK_UPDATE_INDEX As Integer = 25
Public Const UPDATE_COUNT_FOR_ISSUE As Integer = 26

Public Const PUT_POST_NONE As Integer = 0
Public Const PUT_GENERIC_FIELD As Integer = 1
Public Const PUT_GENERIC_UPDATE As Integer = 2
Public Const PUT_GENERIC_UPDATE_REM_ESTIMATE As Integer = 3
Public Const POST_WORKLOG As Integer = 4
Public Const POST_STATUS As Integer = 5

Public Const NO_DESCRIPTION_STRING As String = "No description provided"
Public Const NO_ACCEPTANCE_CRITERIA_STRING As String = "No acceptance criteria provided"
