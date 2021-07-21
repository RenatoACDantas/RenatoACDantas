'@Folder "Planilhas"

Option Explicit

Private Sub Worksheet_Activate()

    ActiveSheet.Unprotect

    Columns("E").ColumnWidth = (Application.Width / 5.88) - (Columns("A").ColumnWidth + Columns("B").ColumnWidth + Columns("C").ColumnWidth + Columns("D").ColumnWidth + Columns("g").ColumnWidth)
    ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:= _
                        False, AllowSorting:=True, AllowFiltering:=True, AllowUsingPivotTables:= _
                        True
End Sub

'Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)

'If Target.Column = 4 And Cells(Target.Row, Target.Column) = "***" Then
'    MsgBox DeCrypt(Cells(Target.Row, 6))
'End If
'End Sub



Private Sub Worksheet_Change(ByVal Target As Range)

    If Target.Column = 4 And Cells(Target.Row, Target.Column) <> "***" Then
        Cells(Target.Row, 6) = Crypt(Cells(Target.Row, Target.Column))
        Cells(Target.Row, Target.Column) = "***"
    End If

End Sub

