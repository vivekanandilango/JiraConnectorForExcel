VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub UserIdSearch_GotFocus()
    UserIdSearch.DropDown
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If track_changes And Target.Count = 1 Then
        ThisWorkbook.Worksheets("ChangeTracker").Range("A1").Value = Target.Value
    End If
End Sub
