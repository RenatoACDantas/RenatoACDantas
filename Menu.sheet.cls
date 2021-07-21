'@Folder "Planilhas"
Option Explicit

Private Sub Worksheet_Activate()

    ActiveSheet.Unprotect
    Columns("D").ColumnWidth = (Application.Width / 5.88) - (Columns("A").ColumnWidth + Columns("B").ColumnWidth + Columns("C").ColumnWidth)
    ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:= _
                        False, AllowSorting:=True, AllowFiltering:=True, AllowUsingPivotTables:= _
                        True
End Sub

