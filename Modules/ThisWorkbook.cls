VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public oldValue As Variant

Private Sub Workbook_Open()
    ThisWorkbook.Worksheets("ChangeTracker").Visible = xlHidden
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    If track_changes And ActiveSheet.Name = SHEET_QUERY_UPDATE Then
        If Selection.Cells.Count = 1 Then
            'On Error GoTo sheet_change_error_handler
            Dim changed_cell_address As String
            Dim table_address As String
            changed_cell_address = Target.Address(0, 0)
            table_address = ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).ListObjects(1).Range.Address(False, False)
            
            On Error Resume Next
            If Application.Intersect(ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range(changed_cell_address), _
                ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range(table_address)) Is Nothing Then
                ' Do nothing
            ElseIf ThisWorkbook.Worksheets("ChangeTracker").Range("A1").Value <> Target.Value Then
                If Target.Value = "" Then
                    ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range(changed_cell_address).Interior.ColorIndex = 3
                Else: ThisWorkbook.Worksheets(SHEET_QUERY_UPDATE).Range(changed_cell_address).Interior.ColorIndex = 6
                End If
            End If
        End If
    End If
'sheet_change_error_handler:
    'EnableEvents
End Sub
