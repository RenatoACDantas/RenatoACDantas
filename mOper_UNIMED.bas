Attribute VB_Name = "mOper_UNIMED"
'@Folder "Operadoras"

Option Explicit


Public mURL As String

Public Sub UNIMED(pLogin As String, pClinica As String, pPass As String, pOperadora As String)
    On Error GoTo unimed_ERR
 
    ' Declarações

    Dim por As New By
    
    Dim mTag As String
    Dim mData
     
    Dim indW As Integer
    Dim indP As Integer
    Dim indDAC As Integer
    Dim ind As Integer
    Dim mIndTag As Integer
    Dim mChData As Integer
    
    Set driver = New ChromeDriver

redo:

    mPasta = CStr(PastaOperadora("Operadora", pOperadora))
    ' mPasta = "C:\Users\renat\teste\UNIMED"
    Debug.Print "PastaOperadora: " & mPasta

    '
    '------------------------------------------------------------------------------------------------------
    '   Propriedades do browser
    '
    driver.setPreference "download.default_directory", mPasta
    driver.setPreference "download.directory_upgrade", True
    driver.setPreference "download.prompt_for_download", False
    
    driver.SetProfile "C:\Users\renat\AppData\Local\Google\Chrome\User Data"
    '
    '------------------------------------------------------------------------------------------------------
    '   Navegar
    '
    If pLogin = "" Or pPass = "" Then GoTo redo
    driver.Start "chrome"
 
    driver.Get "http://autorizador2.unimedse.com.br/autorizador"
    driver.Wait 100

    driver.FindElementById("j_provider").SendKeys pLogin
    driver.FindElementById("j_username").SendKeys pClinica
    driver.FindElementById("j_password_aux").SendKeys pPass
    driver.FindElementById("sub").Click
    driver.Wait 500

    driver.FindElementById("FormMenu:j_id152:j_id256").Click
    driver.Wait 1000
    '
    '-------------------------------------------------------------------------------------------
    '  Navegar para baixar Demonstrativo para Análise de Contas
    '
    
    indW = 2
    indP = 0
    indDAC = 4
    mInd = 0
         
    Sheets("Download").Select
    Columns("A:E").Select
    Selection.ClearContents
    Sheets("Menu").Select
    Range("A1").Select
    '
    '-----------------------------------------------------------------------------------------------
    '   Navegar por parâmetros de datas
    '
    While Sheets("Parametros").Cells(indW, 1) <> ""
        Debug.Print "indw: " & CStr(indW) & " " & mDataInicial
       
        mChData = Sheets("Parametros").Cells(indW, 1)
        mDataInicial = Sheets("Parametros").Cells(indW, 2)
        mDataFinal = Sheets("Parametros").Cells(indW, 3)
        
        Call Navegar
        '
        '---------------------------------------------------------------------------------------------------
        '       Montar Tabela Download
        '
        Dim destino As Range
        Dim tabela As WebElement
        
        Sheets("Download").Select
    
        Set destino = Range("A" & CStr(indDAC + 1))
        Set tabela = driver.FindElementByXPath("//*[@id='Form:table']")
 
        If tabela Is Nothing Then
            MsgBox "Elemento não encontrado"
        Else

            tabela.AsTable.ToExcel destino
            destino.Select
            mInd = Selection.End(xlDown).Row
        End If
    
        Columns("B:B").Select
        Selection.ClearContents
        Range("A1").Select
        '
        '---------------------------------------------------------------------------------------------------
        '   Baixar Demonstratvo para Análise de Contas
        '
        Debug.Print "mInd: " & CStr(mInd)
        For ind = indDAC + 2 To mInd
            Debug.Print "ind: " & CStr(ind)
            Call Navegar
            '
            mIndTag = CStr(ind - (indDAC + 2))
            mTag = Replace("//*[@id='Form:table:0:btnSeeMore']", "0", mIndTag)
            driver.FindElementByXPath(mTag).Click
            driver.Wait 500
        
            driver.FindElementById("Form:btnDown").Click
            driver.Wait 3000
        
        Next ind
    
        indDAC = ind - 1
        
        driver.Wait 1000

        indW = indW + 1

    Wend
    '
    '
    '--------------------------------------------------------------------------------------------------
    '   Fechar browse
    '
    driver.Quit
    '
    '--------------------------------------------------------------------------------------------------
    '   Manipular arquivos baixados
    '
    RenomearArquivos mPasta
    
    'VERIFICAR  RemoverArquivos mPasta
    '
    '
    Exit Sub
    '
unimed_ERR:
    MsgBox Error$
    driver.Quit

End Sub

Sub Navegar()

    Dim por As New By
    
    driver.ExecuteScript ("window.scrollTo(0, 250);")
    driver.FindElementById("FormMenu:j_id152:j_id256").Click
    driver.Wait 1000

    If Not driver.isElementPresent(por.ID("Form:j_id309InputDate")) Then
        driver.FindElementById("iconFormMenu:j_id152:j_id256").Click
        Debug.Print "input nao localizado"
        driver.Wait 1000
    End If
    Debug.Print "mDataInicial: " & mDataInicial
    driver.FindElementByXPath("//*[@id='Form:j_id309InputDate']").Clear
    driver.FindElementByXPath("//*[@id='Form:j_id309InputDate']").SendKeys mDataInicial
    driver.Wait 300
    
    driver.FindElementById("Form:j_id313InputDate").Clear
    driver.FindElementById("Form:j_id313InputDate").SendKeys mDataFinal
    driver.Wait 100

    driver.FindElementById("Form:btnFind").Click
    driver.Wait 1000

End Sub

Sub RenomearArquivos(pPasta As String)

    ' Declarações

    Dim varArquivo As String
    Dim varArquivoNovo As String
    Dim varArquivoAuxiliar As String
    Dim varRow As String
    Dim varArquivosDir As Variant
    Dim varDemonstrativo As String
    Dim varTransacao As String
    Dim varData As String

    Dim temp_1 As String
    Dim temp_2 As String
    Dim temp_3 As String
    
    Dim indW As Integer

    Dim NovoArquivoXLS As Workbook
    Dim sht As Worksheet


    '
    '------------------------------------------------------------------------------------------------
    '   Renomeia DAC
    '

    varArquivosDir = Dir(pPasta & "MedProdTit_*.csv")

    Do While varArquivosDir <> ""
        varArquivo = pPasta & CStr(varArquivosDir)

        Application.DisplayAlerts = False
        Workbooks.Open Filename:=varArquivo, Delimiter:=";"
    
        Application.DisplayAlerts = True
    
        If Mid(CStr(Cells(2, 1)), 1, 8) = "70000025" Then

            temp_1 = InStr(Cells(2, 1), ";")
            temp_2 = Mid(Cells(2, 1), temp_1 + 1)
            temp_3 = InStr(temp_2, ";")

            varData = Mid(temp_2, 1, temp_3 - 1)

        Else
            Debug.Print "Arquivo inválido" & CStr(Cells(1, 2))
        End If

        ActiveWorkbook.Close (False)
    
        indW = 5
        While Sheets("Download").Cells(indW, 1) <> ""
            If CStr(Sheets("Download").Cells(indW, 4)) & "/" & CStr(Sheets("Download").Cells(indW, 3)) = varData Then
            
                varTransacao = Sheets("Download").Cells(indW, 7)
                varDemonstrativo = Sheets("Download").Cells(indW, 3) & Format(Sheets("Download").Cells(indW, 4), "000")
                varArquivoNovo = pPasta & "DAC_" & varDemonstrativo & "_" & varTransacao & ".csv"
                If Dir(varArquivoNovo) <> "" Then
                    Kill varArquivoNovo
                End If
                Name varArquivo As varArquivoNovo

                '
                '           Cria arquivo PAG
                '
            
                varRow = "5:5," & CStr(indW) & ":" & CStr(indW)
                Debug.Print "varRow: " & varRow
                Range(varRow).Select
                Range("A" & CStr(indW)).Activate
                Selection.Copy
            
                varArquivoAuxiliar = pPasta & "PAG_" & varDemonstrativo & "_" & varTransacao & ".xlsx"
                If Dir(varArquivoAuxiliar) <> "" Then
                    Kill varArquivoAuxiliar
                End If
                Set NovoArquivoXLS = Application.Workbooks.Add
                Sheets(1).Select
                Rows("1:1").Select
                ActiveSheet.Paste
            
                NovoArquivoXLS.SaveAs varArquivoAuxiliar
                NovoArquivoXLS.Close (False)
            
                GoTo Saida_DAC
            End If
            indW = indW + 1
    
        Wend
    
Saida_DAC:
        varArquivosDir = Dir(pPasta & "MedProdTit_*.csv")
    
    Loop

    Sheets("Download").Select
    Columns("A:P").Select
    Selection.ClearContents
    Range("A1").Select
    '

End Sub

Sub CriaArquivo(mPlan As Variant, mPathSave As String)

    Dim NovoArquivoXLS As Workbook
    Dim sht As Worksheet

    Set NovoArquivoXLS = Application.Workbooks.Add
    Sheets(1).Select
    Rows("2:2").Select
    Selection.Insert Shift:=xlDown
    NovoArquivoXLS.SaveAs mPathSave
    NovoArquivoXLS.Close (False)

End Sub

Sub RemoverArquivos(pPasta As String)

    Dim mArquivo As String

    mArquivo = Dir(pPasta & "*.xml")
   
    Do While mArquivo <> ""
    
        If Mid(mArquivo, 1, 4) <> "DAC_" And _
                               Mid(mArquivo, 1, 4) <> "PAG_" Then
            Kill pPasta & mArquivo
        End If
        
        mArquivo = Dir
    Loop

End Sub

Sub teste()

    UNIMED "70000025", "70000025", "Digest2021", "UNIMED"

End Sub

Sub d()
    RenomearArquivos "C:\Users\renat\teste\UNIMED\"
End Sub


