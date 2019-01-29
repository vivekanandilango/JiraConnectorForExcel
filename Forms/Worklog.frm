VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Worklog 
   Caption         =   "Worklog"
   ClientHeight    =   11820
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14160
   OleObjectBlob   =   "Worklog.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Worklog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AddUpdateWorklogCommandButton_Click()
    If Worklog.AddUpdateWorklogCommandButton.Caption = "Add" Then
        Call AddWorkLogOfSelectedRow(Worklog.WorklogInput.Value, Replace(Worklog.WorklogInputEditTextBox.Text, vbCrLf, Chr(10)))
        Worklog.WorklogIDs.Clear
        GetWorklog 'Refresh Worklog from JIRA
        If InStr(jira_response, "error") = 0 Then
            Worklog.WorklogInputEditTextBox.Text = ""
        End If
    ElseIf Worklog.AddUpdateWorklogCommandButton.Caption = "Update" Then
        'Add code to update the Worklog to JIRA
        Dim jira_key As String
        jira_key = ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range("A" & ActiveCell.row).Value
        Call SendHttpRequest(API_PUT, jira_key & "/worklog/" & Worklog.WorklogIDs.Value, _
                "{""comment"":""" & Replace(Replace(Worklog.WorklogInputEditTextBox.Text, vbCrLf, Chr(10)), Chr(10), "\n") & _
                """,""timeSpent"": " & Worklog.WorklogInput.Text & "}")
        Worklog.WorklogIDs.Clear
        GetWorklog
        Worklog.AddUpdateWorklogCommandButton.Caption = "Add"
        Worklog.WorklogIDs.Value = ""
        Worklog.WorklogInputEditTextBox.Text = ""
        Worklog.WorklogInput.Text = ""
    End If
End Sub

Private Sub CancelAddCommandButton_Click()
    If Worklog.AddUpdateWorklogCommandButton.Caption = "Add" Then
        Unload Worklog
    Else
        Worklog.AddUpdateWorklogCommandButton.Caption = "Add"
        DeleteCommandButton.Visible = False
        Worklog.WorklogIDs.Value = ""
        Worklog.WorklogInputEditTextBox.Text = ""
        Worklog.WorklogInput.Text = ""
    End If
End Sub

Private Sub WorklogIDs_Change()
    If Worklog.WorklogIDs.Value <> "" Then
        DeleteCommandButton.Visible = True
        Dim worklog_from_jira As String
        worklog_from_jira = ""
        
        Dim all_Worklog As Object
        Set all_Worklog = CallByName(jira_json, "worklogs", VbGet)
        
        Dim current_worklog As Object
        
        Dim Worklog_count As Integer
        Worklog_count = CallByName(jira_json, "total", VbGet)
        
        For counter = 0 To Worklog_count - 1
            'Extract one by one
            Set current_worklog = CallByName(all_Worklog, counter, VbGet)
            If Worklog.WorklogIDs.Value = CallByName(current_worklog, "id", VbGet) Then
                If FieldHasValue(current_worklog, "comment") Then
                    Worklog.WorklogInputEditTextBox.Text = worklog_from_jira & CallByName(current_worklog, "comment", VbGet)
                End If
                Worklog.WorklogInput.Text = CallByName(current_worklog, "timeSpentSeconds", VbGet) / 3600
                Worklog.AddUpdateWorklogCommandButton.Caption = "Update"
                Exit For
            End If
        Next
    ElseIf Worklog.WorklogIDs.Value = "" Then
        Worklog.WorklogInputEditTextBox.Text = ""
        Worklog.AddUpdateWorklogCommandButton.Caption = "Add"
    End If
End Sub

Private Sub DeleteCommandButton_Click()
    Dim jira_key As String
    jira_key = ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range("A" & ActiveCell.row).Value
    Call SendHttpRequest("DELETE", jira_key & "/worklog/" & Worklog.WorklogIDs.Value, "")
    Worklog.WorklogIDs.Clear
    DeleteCommandButton.Visible = False
    GetWorklog
    Worklog.AddUpdateWorklogCommandButton.Caption = "Add"
    Worklog.WorklogIDs.Value = ""
    Worklog.WorklogInputEditTextBox.Text = ""
    Worklog.WorklogInput.Text = ""
End Sub

'Handle escape key

Private Sub AddUpdateWorklogCommandButton_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then
        Unload Worklog
    End If
End Sub

Private Sub CancelAddCommandButton_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then
        Unload Worklog
    End If
End Sub

Private Sub DeleteCommandButton_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then
        Unload Worklog
    End If
End Sub

Private Sub UserForm_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then
        Unload Worklog
    End If
End Sub

Private Sub WorklogIDs_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then
        Unload Worklog
    End If
End Sub

Private Sub WorklogInput_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then
        Unload Worklog
    End If
End Sub

Private Sub WorklogInputEditTextBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then
        Unload Worklog
    End If
End Sub

Private Sub WorklogText_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then
        Unload Worklog
    End If
End Sub
