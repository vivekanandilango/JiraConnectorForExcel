Attribute VB_Name = "FieldList"
Sub GetFields()
    GetHttpRequest ("https://<jira-url>/rest/api/latest/issue/createmeta?projectKeys=KEY&issuetypeNames=Type&expand=projects.issuetypes.fields")
    Dim projects As Variant
    Dim project_name As Variant
    Set project_name = CallByName(CallByName(jira_json, "projects", VbGet), 0, VbGet)
    Dim project_issue_fields As Variant
    Set project_issue_fields = CallByName(CallByName(CallByName(project_name, "issuetypes", VbGet), 0, VbGet), "fields", VbGet)
    Dim field_length As Integer
    field_length = Len(project_issue_fields)
    Dim project_count As Integer
    project_count = CallByName(projects, "length", VbGet)
    For counter = 0 To project_count - 1
        If CallByName(CallByName(projects, counter, VbGet), "key", VbGet) = "KEY" Then
            Set project_name = CallByName(projects, counter, VbGet)
            Exit For
        End If
    Next
End Sub
