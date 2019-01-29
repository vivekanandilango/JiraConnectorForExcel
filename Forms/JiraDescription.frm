VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} JiraDescription 
   Caption         =   "Description"
   ClientHeight    =   9015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14160
   OleObjectBlob   =   "JiraDescription.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "JiraDescription"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CancelCommandButton_Click()
    If JiraDescription.ModifyCommandButton.Caption = "Modify" Then
        'Default cancel
        Unload JiraDescription
    ElseIf JiraDescription.ModifyCommandButton.Caption = "Update" Then
        'If update was selected, need to cancel the update and go back to initial state
        JiraDescription.ModifyCommandButton.Caption = "Modify"
        JiraDescription.DescriptionText.BackColor = &HE0E0E0
        JiraDescription.DescriptionText.Locked = True
    End If
End Sub

Private Sub ModifyCommandButton_Click()
    If JiraDescription.ModifyCommandButton.Caption = "Modify" Then
        'Enable provision for user to modify the existing description
        JiraDescription.ModifyCommandButton.Caption = "Update"
        If JiraDescription.DescriptionText.Value = NO_DESCRIPTION_STRING Then
            'Clear the textbox when no description provided
            JiraDescription.DescriptionText.Value = ""
        End If
        JiraDescription.DescriptionText.BackColor = &H80000005
        JiraDescription.DescriptionText.Locked = False
    Else
        'Update the description to JIRA
        Dim jira_key As String
        jira_key = ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range("A" & ActiveCell.row).Value 'Get jira key
        Call SendHttpRequest(API_PUT, jira_key, _
                "{""fields"":{""description"":""" & Replace(Replace(JiraDescription.DescriptionText.Value, vbCrLf, Chr(10)), Chr(10), "\n") & """}}")
        GetDescription 'Refresh description field
        JiraDescription.ModifyCommandButton.Caption = "Modify"
        JiraDescription.DescriptionText.BackColor = &HC0C0C0
        JiraDescription.DescriptionText.Locked = True
    End If
End Sub

'Handle escape key

Private Sub ModifyCommandButton_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then
        Unload JiraDescription
    End If
End Sub

Private Sub UserForm_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then
        Unload JiraDescription
    End If
End Sub

Private Sub CancelCommandButton_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then
        Unload JiraDescription
    End If
End Sub

Private Sub DescriptionText_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then
        Unload JiraDescription
    End If
End Sub
