VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Comments 
   Caption         =   "Comments"
   ClientHeight    =   11820
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14160
   OleObjectBlob   =   "Comments.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Comments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AddUpdateCommentCommandButton_Click()
    If Comments.AddUpdateCommentCommandButton.Caption = "Add" And Comments.CommentInputEditTextBox.Text <> "" Then
        AddCommentOfSelectedRow (Replace(Comments.CommentInputEditTextBox.Text, vbCrLf, Chr(10)))
        Comments.CommentIDs.Clear
        GetComment 'Refresh comments from JIRA
        If InStr(jira_response, "error") = 0 Then
            Comments.CommentInputEditTextBox.Text = ""
        End If
    ElseIf Comments.AddUpdateCommentCommandButton.Caption = "Update" Then
        Dim jira_key As String
        jira_key = ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range("A" & ActiveCell.row).Value
        Call SendHttpRequest(API_PUT, jira_key & "/comment/" & Comments.CommentIDs.Value, _
                "{""body"":""" & Replace(Replace(Comments.CommentInputEditTextBox.Text, vbCrLf, Chr(10)), Chr(10), "\n") & """}")
        Comments.CommentIDs.Clear
        GetComment
        Comments.AddUpdateCommentCommandButton.Caption = "Add"
        Comments.CommentIDs.Value = ""
        Comments.CommentInputEditTextBox.Text = ""
    End If
End Sub

Private Sub CancelAddCommandButton_Click()
    If Comments.AddUpdateCommentCommandButton.Caption = "Add" Then
        Unload Comments
    Else
        Comments.AddUpdateCommentCommandButton.Caption = "Add"
        DeleteCommandButton.Visible = False
        Comments.CommentIDs.Value = ""
        Comments.CommentInputEditTextBox.Text = ""
    End If
End Sub

Private Sub CommentIDs_Change()
    If Comments.CommentIDs.Value <> "" Then
        DeleteCommandButton.Visible = True
        Dim comments_from_jira As String
        comments_from_jira = ""
        
        Dim all_comments As Object
        Set all_comments = CallByName(jira_json, "comments", VbGet)
        
        Dim current_comment As Object
        
        Dim comment_count As Integer
        comment_count = CallByName(jira_json, "total", VbGet)
        
        For counter = 0 To comment_count - 1
            'Extract one by one
            Set current_comment = CallByName(all_comments, counter, VbGet)
            'Extract date and time from comment date
            If Comments.CommentIDs.Value = CallByName(current_comment, "id", VbGet) Then
                Comments.CommentInputEditTextBox.Text = comments_from_jira & CallByName(current_comment, "body", VbGet)
                Comments.AddUpdateCommentCommandButton.Caption = "Update"
                Exit For
            End If
        Next
    ElseIf Comments.CommentIDs.Value = "" Then
        Comments.CommentInputEditTextBox.Text = ""
        Comments.AddUpdateCommentCommandButton.Caption = "Add"
    End If
End Sub

Private Sub DeleteCommandButton_Click()
    Dim jira_key As String
    jira_key = ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range("A" & ActiveCell.row).Value
    Call SendHttpRequest("DELETE", jira_key & "/comment/" & Comments.CommentIDs.Value, "")
    Comments.CommentIDs.Clear
    DeleteCommandButton.Visible = False
    GetComment
    Comments.AddUpdateCommentCommandButton.Caption = "Add"
    Comments.CommentIDs.Value = ""
    Comments.CommentInputEditTextBox.Text = ""
End Sub

'Handle escape key
Private Sub UserForm_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then
        Unload Comments
    End If
End Sub

Private Sub AddUpdateCommentCommandButton_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then
        Unload Comments
    End If
End Sub

Private Sub CancelAddCommandButton_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then
        Unload Comments
    End If
End Sub

Private Sub CommentIDs_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then
        Unload Comments
    End If
End Sub

Private Sub CommentInputEditTextBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then
        Unload Comments
    End If
End Sub

Private Sub CommentText_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then
        Unload Comments
    End If
End Sub

Private Sub DeleteCommandButton_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then
        Unload Comments
    End If
End Sub
