Attribute VB_Name = "mOper_CASSI"
'@Folder "Operadoras"
Option Explicit

Public MyPass As String
Public MyLogin As String
Public mDataInicial As String
Public mDataFinal As String
Public mSeqData As Integer


Public driver As ChromeDriver



'
'----------------------------------------------------------------------------------------------------
'   SCRIPT CASSI
'----------------------------------------------------------------------------------------------------
'

Public Sub CASSI(pLogin As String, pPass As String, pOperadora As String)
    On Error GoTo cassi_ERR

    ' Declarações

    Dim por As New By
    
    Dim ele As Variant
    
    Dim mProtocolo As String
    Dim mDataAnterior As String
    
    
    Dim indW As Integer
    Dim indP As Integer
    Dim indDAC As Integer
    Dim mInd As Integer
    Dim ind As Integer
    Dim mChData As Integer
    
 
    Set driver = New ChromeDriver
    

redo:

    mPasta = CStr(PastaOperadora("Operadora", pOperadora))
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
 
    driver.Get "https://www.cassi.com.br/index.php?option=com_content&view=featured&Itemid=344&uf=&uf=SE"
    driver.Wait 1000

    driver.Get "https://servicosonline.cassi.com.br/GASC/v2/Prestador"
    

    '   driver.findElementById("cpfcnpj").SendKeys pLogin
    '   driver.findElementById("Senha").SendKeys pPass
    driver.FindElementById("btnSubmitSemAjax").Click
    driver.Wait 500

    driver.Get "https://servicosonline.cassi.com.br/GASC/v2/Usuario/MeusDados"
    driver.Wait 1000
    '
    '-------------------------------------------------------------------------------------------
    '  Navegar para baixar Demonstrativo para Análise de Contas
    '
    
    indW = 2
    indP = 0
    indDAC = 0
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
        
        mChData = Sheets("Parametros").Cells(indW, 1)
        mDataInicial = Sheets("Parametros").Cells(indW, 2)
        mDataFinal = Sheets("Parametros").Cells(indW, 3)
        
        driver.Get "https://servicosonline.cassi.com.br/Prestador/RecursoRevisaoPagamento/TISS/DemonstrativoAnaliseContas/Index"
       
        driver.FindElementById("DataInicial").SendKeys mDataInicial
        driver.FindElementById("DataFinal").Click
    
        driver.FindElementById("DataFinal").SendKeys mDataFinal
        driver.Wait 1000
    
        driver.FindElementByName("ProtocoloPagamento").Click
        driver.Wait 1000
    
        driver.FindElementById("btnConsultar").Click
        driver.Wait 1000

        '
        '---------------------------------------------------------------------------------------------------
        '   Baixar Demonstratvo para Análise de Contas
        '
        mInd = indDAC
        For Each ele In driver.FindElementsByCss(".table tbody tr td:nth-child(1)")
            mInd = mInd + 1
            Sheets("Download").Cells(mInd, 1) = mChData
            Sheets("Download").Cells(mInd, 2) = ele.Text
        Next
        
        mInd = indDAC
        For Each ele In driver.FindElementsByCss(".table tbody tr td:nth-child(2)")
            mInd = mInd + 1
            Sheets("Download").Cells(mInd, 3) = CDate(ele.Text)
        Next

        For ind = indDAC + 1 To mInd

            driver.Get "https://servicosonline.cassi.com.br/Prestador/RecursoRevisaoPagamento/TISS/DemonstrativoAnaliseContas/Index"
            driver.Wait 100
    
            mProtocolo = Sheets("Download").Cells(ind, 2)
            '--            Debug.Print "Protocolo: " & Sheets("Download").Cells(ind, 2)
            driver.FindElementByName("ProtocoloPagamento").SendKeys mProtocolo
            driver.Wait 100
    
            driver.FindElementById("btnConsultar").Click
            driver.Wait 1000

            driver.FindElementByXPath("/html/body/div[1]/div[5]/section/div/fieldset/div/table/tbody/tr/td[3]/form/input[3]").Click
            driver.Wait 1000
            '--            Debug.Print CStr(ind) + " " + CStr(indP) + " " + CStr(mInd) + " " + CStr(Sheets("Download").Cells(ind, 3))
    
            driver.FindElementByXPath("/html/body/div[1]/div[5]/section/div/div[1]/div[1]/div/div[1]/form/button[2]").Click
            driver.Wait 100

        Next ind
        
        indDAC = ind - 1
        
        driver.Wait 10000
        '
        '---------------------------------------------------------------------------------------------------
        '   Baixar Demonstrativo de Pagamento
        '

        driver.Get "https://servicosonline.cassi.com.br/Prestador/RecursoRevisaoPagamento/TISS/DemonstrativoPagamento/Index"

        driver.FindElementById("DataInicial").SendKeys mDataInicial
        driver.FindElementById("DataFinal").Click
    
        driver.FindElementById("DataFinal").SendKeys mDataFinal
        driver.Wait 100
                
        '--     Debug.Print "XML Pagamento: " & mDataFinal
        driver.FindElementById("btnConsultar").Click
        driver.Wait 1000


        mInd = indP
        For Each ele In driver.FindElementsByCss(".table tbody tr td:nth-child(1)")
            mInd = mInd + 1
            Sheets("Download").Cells(mInd, 5) = mChData
            Sheets("Download").Cells(mInd, 6) = CDate(ele.Text)
        Next
        
        For ind = indP + 1 To mInd
            '--            Debug.Print CStr(ind) & " " & CStr(Sheets("Download").Cells(ind, 3)) & " " & mDataInicial
            If Sheets("Download").Cells(ind, 6) <> mDataAnterior Then
                driver.Get "https://servicosonline.cassi.com.br/Prestador/RecursoRevisaoPagamento/TISS/DemonstrativoPagamento/Index"

                mDataInicial = Sheets("Download").Cells(ind, 6)
                mDataFinal = mDataInicial
                driver.FindElementById("DataInicial").SendKeys mDataInicial
                driver.FindElementById("DataFinal").Click
    
                driver.FindElementById("DataFinal").SendKeys mDataFinal
                driver.Wait 100
                
                driver.FindElementById("btnConsultar").Click
                driver.Wait 1000
    
                driver.FindElementByXPath("/html/body/div[1]/div[5]/section/div/fieldset/div/table/tbody/tr[1]/td[2]/form/input[2]").Click
                driver.Wait 1000
 
                If driver.isElementPresent(por.XPath("//*[@id=""formExportar""]/button[2]")) = True Then
                    driver.FindElementByXPath("//*[@id=""formExportar""]/button[2]").Click
                    driver.Wait 1000
                    Debug.Print "pag-2"
                Else
                    Debug.Print "PAG: Data sem arquivo " & mDataInicial
                End If
                Debug.Print "pag-3"
            End If
        
            mDataAnterior = Sheets("Download").Cells(ind, 6)
        Next ind
        
        indP = ind - 1
        
        driver.Wait 10000

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
    
    RemoverArquivos mPasta
    '
    '
    Exit Sub
    '
cassi_ERR:
    MsgBox Error$
    driver.Quit

End Sub

Sub RenomearArquivos(pPasta As String)

    ' Declarações

    Dim varArquivoNovo As String
    Dim varArquivo As String
    Dim mDTAnterior As String
    Dim varDemonstrativo As String

    Dim ind As Integer

    ind = 1
    mDTAnterior = ""


    '
    '   Arquivo DAC
    '
    ind = 1

    Do While Sheets("Download").Cells(ind, 1) <> ""

        varArquivo = pPasta & Sheets("Download").Cells(ind, 2) & ".xml"
        If Dir(varArquivo) <> vbNullString Then
            varArquivoNovo = pPasta & _
                             "DAC_" & Format(Sheets("Download").Cells(ind, 3), "yyyymmdd") & "_" & Sheets("Download").Cells(ind, 2) & ".xml"

            If Dir(varArquivoNovo) <> vbNullString Then
                Kill varArquivoNovo
            End If
            Name varArquivo As varArquivoNovo
        
            Sheets("Download").Cells(ind, 4) = "Sim"
        Else
            Debug.Print "nao: " & varArquivo
            Sheets("Download").Cells(ind, 4) = "Não"
        End If
    
        ind = ind + 1

    Loop

    '
    '   Arquivo PAG
    '
    ind = 1
    mDTAnterior = ""

    Do While Sheets("Download").Cells(ind, 5) <> ""

        varArquivo = pPasta & Format(Sheets("Download").Cells(ind, 6), "dmyyyy") & ".xml"
        If Sheets("Download").Cells(ind, 6) <> mDTAnterior Then
            If Dir(varArquivo) <> vbNullString Then
        
                Application.DisplayAlerts = False
                Workbooks.OpenXML varArquivo, , xlXmlLoadImportToList
                Application.DisplayAlerts = True
                If CStr(Cells(1, 9)) = "ns1:numeroDemonstrativo" Then
                    varDemonstrativo = CStr(Cells(2, 9))
                Else
                    varDemonstrativo = CStr(Cells(1, 9))
                End If
                ActiveWorkbook.Close (False)
               
                varArquivoNovo = pPasta & _
                                 "PAG_" & Format(Sheets("Download").Cells(ind, 6), "yyyymmdd") & "_" & varDemonstrativo & ".xml"
                If Dir(varArquivoNovo) <> vbNullString Then
                    Kill varArquivoNovo
                End If

                Name varArquivo As varArquivoNovo
            
                Sheets("Download").Cells(ind, 7) = "Sim"
            Else
                Debug.Print "nao: " & varArquivo
                Sheets("Download").Cells(ind, 7) = "Não"
            End If
        End If
        mDTAnterior = Sheets("Download").Cells(ind, 6)
    
        ind = ind + 1

    Loop

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

Public Function WaitNewFile(Optional Target As String) As String
    Static files As Collection, filter$
    Dim file$, file_path$, i&
    If Len(Target) Then
        ' Initialize the list of files and return
        filter = Target
        Set files = New Collection
        file = Dir(filter, vbNormal)
        Do While Len(file)
            files.Add Empty, file
            file = Dir
        Loop
        Exit Function
    End If
  
    ' Waits for a file that is not in the list
    On Error GoTo WaitReady
    Do
        file = Dir(filter, vbNormal)
        Do While Len(file)
            files.Item file
            file = Dir
        Loop
        For i = 0 To 3000: DoEvents: Next
    Loop
  
WaitReady:
    ' Waits for the size to be superior to 0 and try to rename it
    file_path = Left$(filter, InStrRev(filter, "\")) & file
    Do
        If FileLen(file_path) Then
            On Error Resume Next
            Name file_path As file_path
            If Err = 0 Then Exit Do
        End If
        For i = 0 To 3000: DoEvents: Next
    Loop
    files.Add Empty, file
    WaitNewFile = file_path
End Function

'Função que identifica a existência do arquivo
Public Function lfVerificaArquivo(ByVal lStr As String) As Boolean

    lfVerificaArquivo = True
    
    'Identifica se o arquivo existe
    If Dir(lStr) = vbNullString Then
        lfVerificaArquivo = False
        Debug.Print "O arquivo: '" & lStr & "' não foi encontrado! Por favor verifique o caminho e a extensão do arquivo"
    Else
        lfVerificaArquivo = True
    End If
    
End Function

Sub teste()
    RenomearArquivos "C:\Users\renat\Downloads\health TEC\digest\CASSI\"
End Sub


