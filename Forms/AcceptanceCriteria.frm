VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AcceptanceCriteria 
   Caption         =   "Acceptance Criteria"
   ClientHeight    =   9015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14160
   OleObjectBlob   =   "AcceptanceCriteria.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AcceptanceCriteria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CancelCommandButton_Click()
    If AcceptanceCriteria.ModifyCommandButton.Caption = "Modify" Then
        'Default cancel
        Unload AcceptanceCriteria
    ElseIf AcceptanceCriteria.ModifyCommandButton.Caption = "Update" Then
        'If update was selected, need to cancel the update and go back to initial state
        AcceptanceCriteria.ModifyCommandButton.Caption = "Modify"
        AcceptanceCriteria.AcceptanceCriteriaText.BackColor = &HC0C0C0
        AcceptanceCriteria.AcceptanceCriteriaText.Locked = True
    End If
End Sub

Private Sub ModifyCommandButton_Click()
    If AcceptanceCriteria.ModifyCommandButton.Caption = "Modify" Then
        'Enable provision for user to modify the existing AcceptanceCriteria
        AcceptanceCriteria.ModifyCommandButton.Caption = "Update"
        If AcceptanceCriteria.AcceptanceCriteriaText.Value = NO_ACCEPTANCE_CRITERIA_STRING Then
            'Clear the textbox when no AcceptanceCriteria provided
            AcceptanceCriteria.AcceptanceCriteriaText.Value = ""
        End If
        AcceptanceCriteria.AcceptanceCriteriaText.BackColor = &H80000005
        AcceptanceCriteria.AcceptanceCriteriaText.Locked = False
    Else
        'Update the AcceptanceCriteria to JIRA
        Dim jira_key As String
        jira_key = ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range("A" & ActiveCell.row).Value 'Get jira key
        Call SendHttpRequest(API_PUT, jira_key, _
                "{""fields"":{""customfield_xxxxx"":""" & Replace(Replace(AcceptanceCriteria.AcceptanceCriteriaText.Value, vbCrLf, Chr(10)), Chr(10), "\n") & """}}")
        GetAcceptanceCriteria 'Refresh AcceptanceCriteria field
        AcceptanceCriteria.ModifyCommandButton.Caption = "Modify"
        AcceptanceCriteria.AcceptanceCriteriaText.BackColor = &HC0C0C0
        AcceptanceCriteria.AcceptanceCriteriaText.Locked = True
    End If
End Sub

'Handle escape key

Private Sub ModifyCommandButton_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then
        Unload AcceptanceCriteria
    End If
End Sub

Private Sub UserForm_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then
        Unload AcceptanceCriteria
    End If
End Sub

Private Sub CancelCommandButton_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then
        Unload AcceptanceCriteria
    End If
End Sub

Private Sub AcceptanceCriteriaText_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then
        Unload AcceptanceCriteria
    End If
End Sub


