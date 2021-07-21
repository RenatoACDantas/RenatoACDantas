'@Folder "Planilhas"
Option Explicit

Private Sub Worksheet_Activate()


    ActiveSheet.Unprotect
    Columns("D").ColumnWidth = (Application.Width / 5.88) - (Columns("A").ColumnWidth + Columns("B").ColumnWidth + Columns("C").ColumnWidth)
    ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:= _
                        False, AllowSorting:=True, AllowFiltering:=True, AllowUsingPivotTables:= _
                        True

End Sub

Private Sub Worksheet_Change(ByVal Target As Range)

    If Cells(Target.Row, 2) = "Geral" And Cells(Target.Row, 3) = "Pasta Padrão" Then
        Debug.Print Cells(Target.Row, 4)
        Alterar_PastaPadrão (Cells(Target.Row, 4))
    End If

    'Pasta Padrão


End Sub

