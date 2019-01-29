VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SessionIdQuery 
   Caption         =   "Session ID Query"
   ClientHeight    =   2265
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8430
   OleObjectBlob   =   "SessionIdQuery.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SessionIdQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub GetSessionId_Click()
    ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range("D1").Value = ""
    Dim JiraAuth As New MSXML2.XMLHTTP60
    With JiraAuth
        .Open "POST", JIRA_API_AUTH_URL, False
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "X-Atlassian-Token", "nocheck"
        .setRequestHeader "Accept", "application/json"
        .send " {""username"" : """ & SessionIdQuery.UserIdInput.Value & """ , ""password"" : """ & SessionIdQuery.PasswordInput.Value & """}"
        responseText = .responseText
        If .status = 200 Then
            sessionId = "JSESSIONID=" & Mid(responseText, 42, 32) & "; Path=/Jira"
            ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range("D1").Value = sessionId
            Unload SessionIdQuery
            .abort
            ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range("B7").Value = DateTime.Now
            ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range("B7").NumberFormat = "hh:mm:ss AM/PM"
        Else
            MsgBox "Authentication Failed, Exiting."
            Unload SessionIdQuery
            End
        End If
    End With
End Sub

'This should handle if user hits enter from password box
'Caveat, on every selection of the button it will trigger button click
Private Sub GetSessionId_Enter()
    GetSessionId_Click
End Sub
