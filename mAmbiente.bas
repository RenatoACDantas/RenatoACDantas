Attribute VB_Name = "mAmbiente"
'@Folder "Ambiente"
Option Explicit

Sub usuario()

    ' Declarações

    Dim Computador As String
    Dim dominio As String
    Dim user As String

    Computador = Environ("Computername")
    dominio = Environ("USERDOMAIN")
    user = Environ("USERNAME")

    ' Range("a1") = user

End Sub

Sub Formula()
    '
    Range("C8").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(R3C[1],Dados!C[-2]:C[8],8,0),"""")"

    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(R3C[1],Dados!C[-2]:C[8],8,0),"""")"
    Range("c9").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(R3C[1],Dados!C[-2]:C[8],6,0),"""")"
    Range("c10").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(R3C[1],Dados!C[-2]:C[8],7,0),"""")"
    Range("c11").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(R3C[1],Dados!C[-2]:C[8],9,0),"""")"
    
    Range("c8").Select
    
End Sub

Sub Sair()
    Dim barras
    On Error Resume Next
    For Each barras In Application.CommandBars
        barras.Enabled = True
    Next
    Application.DisplayStatusBar = True
    Application.DisplayFormulaBar = True
    Application.DisplayFullScreen = False
    ActiveWindow.DisplayHeadings = True
    ActiveWindow.DisplayHorizontalScrollBar = True
    ActiveWindow.DisplayVerticalScrollBar = True
    ActiveWindow.DisplayWorkbookTabs = True
End Sub

Sub Desabilitar()
    On Error Resume Next

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    '   Application.EnableEvents = False
      
    Application.DisplayStatusBar = False
    Application.DisplayFormulaBar = False
    Application.DisplayFullScreen = False
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.DisplayHorizontalScrollBar = False
    ActiveWindow.DisplayVerticalScrollBar = False
    ActiveWindow.DisplayWorkbookTabs = False
End Sub

Sub Habilitar()
    On Error Resume Next

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    Application.DisplayStatusBar = True
    Application.DisplayFormulaBar = True
    Application.DisplayFullScreen = False
    ActiveWindow.DisplayHeadings = True
    ActiveWindow.DisplayHorizontalScrollBar = True
    ActiveWindow.DisplayVerticalScrollBar = True
    ActiveWindow.DisplayWorkbookTabs = True
End Sub

Sub CloseBook()
    Application.DisplayAlerts = False
    ActiveWorkbook.Close savechanges:=False
    Application.DisplayAlerts = True
    
End Sub

Sub CriarFolder(pPasta As String)
    On Error Resume Next
    If Dir(pPasta) = "" Then
        MkDir (pPasta)
    End If
    
End Sub

Function TaskManager(pProcess As String) As Integer
    Dim oServ As Object
    Dim cProc As Variant
    Dim oProc As Object
    Dim mResultado As Integer
Debug.Print "tm:" & pProcess
    Set oServ = GetObject("winmgmts:")
    Set cProc = oServ.ExecQuery("Select * from Win32_Process")
    mResultado = 0
    
    For Each oProc In cProc
        If oProc.Name = "chrome.exe" Then
            mResultado = MsgBox("Google Chrome em execução. Será necessário encerrar.", vbOKCancel)
            If mResultado = 1 Then
                killProcess pProcess
                
            End If
            GoTo sai
        End If
    Next
    
sai:
    TaskManager = mResultado
    
End Function

Sub killProcess(pProcess As String)

    Dim Process As String
    Debug.Print pProcess
    
    Process = "TASKKILL /F /IM " & pProcess
    Shell Process, vbHide
    
    Application.Wait Now + TimeValue("00:00:05")
    
End Sub


