Attribute VB_Name = "mOper_CASSE"
'@Folder "Operadoras"
Option Explicit


Public mURL As String

Public Sub CASSE(pLogin As String, pPass As String, pOperadora As String)
    On Error GoTo casse_ERR
 
    ' Declarações

    Dim por As New By
    
    Dim ele As Variant
    
    Dim mChData As Integer
    Dim indX As Integer
    Dim indW As Integer
    Dim indDAC As Integer
    Dim indDown As Integer
    Dim ind As Integer
 
 
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
    '   Login
    '
    '    If MyLogin = "" Or MyPass = "" Then GoTo redo
    '    driver.Start "chrome"
 
    driver.Get "https://casse.banese.com.br/portalcasse/prestador/index.php"
    driver.Wait 500

    '    driver.FindElementById("operador").SendKeys pLogin
    '    driver.FindElementById("senha").SendKeys pPass
    
    driver.Wait 10000

    '
    '-------------------------------------------------------------------------------------------------------------------
    '   captcha
    '
    '   driver.switchToFrame (0)
    '   driver.findElementByXPath("/html/body/div[2]/div[3]/div[1]/div/div/span/div[4]").SendKeys True

    '   driver.SwitchToFrame.FindElementByXPath("//iframe[contains(@src, ""recaptcha"") and not(@title=""recaptcha challenge"")]", timeout:=10000)
    '   driver.FindElementByCss("div.recaptcha-checkbox-checkmark").Click
    '   driver.findElementById("entrar").Click
    '
    '----------------------------------------------------------------------------------------------------------------------
    '

    '
    '------------------------------------------------------------------------------------------------------
    '   Navega (e aguarda resposta manual de captcha
    '
    indX = 0
    Debug.Print "1: "
    While Not driver.isElementPresent(por.Css("#menu_nav > div > div.container > div.span15 > div > ul > li:nth-child(5) > a")) And indX < 50
        Application.Wait Now + TimeValue("00:00:03")
        driver.Wait 1000
        indX = indX + 1
    Wend
    If indX >= 50 Then
        Err.Description = "site lento. Elemento não alcançado: li.dropdown.open)"
        Err.Number = 999
        GoTo casse_ERR
    End If
    driver.Wait 5000
    
    Debug.Print "2: "
    driver.FindElementByCss("#menu_nav > div > div.container > div.span15 > div > ul > li:nth-child(5) > a").Click
    driver.Wait 1000
    
    indX = 0
    Debug.Print "3: "
    While Not driver.isElementPresent(por.Css("#menu_nav > div > div.container > div.span15 > div > ul > li.dropdown.open > ul > li.dropdown-submenu.menuAcess > a")) And indX < 10
        Debug.Print "3: " & CStr(indX) & " " & CStr(Now)
        Application.Wait Now + TimeValue("00:00:01")
        driver.Wait 1000
        indX = indX + 1
    Wend
    If indX >= 10 Then
        Err.Description = "site lento. Elemento não alcançado: li.DropDown - submenu.menuAcess"
        Err.Number = 999
        GoTo casse_ERR
    End If
    
    Debug.Print "4: "
    driver.FindElementByCss("#menu_nav > div > div.container > div.span15 > div > ul > li.dropdown.open > ul > li.dropdown-submenu.menuAcess > a").ReleaseMouse
    driver.Wait 500
    
    Debug.Print "5: "
    driver.FindElementByCss("#menu_nav > div > div.container > div.span15 > div > ul > li.dropdown.open > ul > li.dropdown-submenu.menuAcess > ul > li:nth-child(1) > a").Click
    driver.Wait 500
    '
    '------------------------------------------------------------------------------------------------------
    '   Filtra por data
    '
    
    indW = 2
    indDAC = 0
    indDown = 0
    mInd = 0
    Debug.Print "download"
    Sheets("Download").Select
    Columns("A:E").Select
    Selection.ClearContents
    '   Sheets("Menu").Select
    '   Range("A1").Select
    '
    '-----------------------------------------------------------------------------------------------
    '   Navegar por parâmetros de datas
    '
    
    Debug.Print "6: "
    While Sheets("Parametros").Cells(indW, 1) <> ""
        Debug.Print "Data: " & CStr(indW) & CStr(Sheets("Parametros").Cells(indW, 2))
        mChData = Sheets("Parametros").Cells(indW, 1)
        mDataInicial = Replace(Sheets("Parametros").Cells(indW, 2), "/", "")
        mDataFinal = Replace(Sheets("Parametros").Cells(indW, 3), "/", "")

        driver.FindElementById("data_ini").Clear
        driver.FindElementById("data_ini").SendKeys mDataInicial
        driver.Wait 100

        driver.FindElementById("data_fim").Clear
        driver.FindElementById("data_fim").SendKeys mDataFinal
        driver.Wait 1000
        driver.ExecuteScript ("window.scrollTo(0,document.body.scrollHeight);")
            
        driver.FindElementById("enviar").Click
        driver.Wait 1000
        Debug.Print driver.Url
        mURL = driver.Url
  
        '
        '-------------------------------------------------------------------------------------------
        '  Navegar para baixar Demonstrativos
        '
     
        Debug.Print "7: " & CStr(indDAC)
        indDown = 0
        mInd = indDAC
        For Each ele In driver.FindElementsByCss("#meio > div.container > table > tbody > tr > td:nth-child(2) > a")
            mInd = mInd + 1
            indDown = indDown + 1
            Sheets("Download").Cells(mInd, 3) = CDate(ele.Text)
        Next
  
        mInd = indDAC
        For Each ele In driver.FindElementsByCss("#meio > div.container > table > tbody > tr > td:nth-child(3) > a")
            mInd = mInd + 1
            Sheets("Download").Cells(mInd, 2) = ele.Text
        Next
        indDAC = mInd
        '
        '---------------------------------------------------------------------------------------------------
        '   Baixar Demonstratvo para Análise de Contas
        '
    
        Debug.Print "8: " & CStr(indDown)
        For ind = 1 To indDown
            Debug.Print "8.1: " & CStr(ind)
            driver.FindElementByXPath("//*[@id=""meio""]/div[1]/table/tbody/tr[" + CStr(ind) + "]/td[15]/a/img").Click
            driver.Wait 1000
            '
            '---------------------------------------------------------------------------------------------------
            '   Baixar Demonstrativo de Pagamento
            '
            Debug.Print "8.2: " & CStr(ind)
            driver.FindElementByXPath("//*[@id=""meio""]/div[1]/table/tbody/tr[" + CStr(ind) + "]/td[16]/a/img").Click
            driver.Wait 1000

        Next ind
  
        driver.Wait 500
        indW = indW + 1
        driver.Refresh
        driver.Get mURL
        driver.Wait 1000
    Wend
    '
    '--------------------------------------------------------------------------------------------------
    '   Fechar browse
    '
    driver.Quit
    '
    '--------------------------------------------------------------------------------------------------
    '   Manipular arquivos baixados
    '
    Unzip mPasta
    RenomearArquivos mPasta

    
    
    
    Exit Sub
    '
casse_ERR:
    Debug.Print Err.Description
    Debug.Print Error$
    MsgBox Error$
    driver.Quit

End Sub

Sub Unzip(pPasta As String)

    Dim oApp As Object
    Dim varArquivo As Variant
    Dim varArquivosDir As Variant
    Dim Caminho As String
    Dim NovaPasta As Variant

    NovaPasta = CStr(pPasta)

    varArquivosDir = Dir(pPasta & "*.zip")
    Debug.Print pPasta

    Do While varArquivosDir <> ""
        Debug.Print varArquivosDir
        varArquivo = pPasta & CStr(varArquivosDir)
        Debug.Print varArquivo
        '
        '------------------------------------------------------------------------------------------------
        '   Extrai os arquivos para a pasta informada
        '
        Set oApp = CreateObject("Shell.Application")
        oApp.Namespace(NovaPasta).CopyHere oApp.Namespace(varArquivo).Items
        '
        '   Remove arquivo Zip
        '
        Kill varArquivo
    
        varArquivosDir = Dir()
    Loop

End Sub

Sub RenomearArquivos(pPasta As String)

    ' Declarações
    
    Dim varArquivo As String
    Dim varArquivoNovo As String
    Dim varArquivosDir As Variant
    Dim varDemonstrativo As String
    Dim varTransacao As String
    Dim varData As String
    
    Dim indW As Integer
    '
    '------------------------------------------------------------------------------------------------
    '   Renomeia DAC
    '
    varArquivosDir = Dir(pPasta & "demonstrativo_*.xml")

    Do While varArquivosDir <> ""
        varArquivo = pPasta & CStr(varArquivosDir)
    
        Application.DisplayAlerts = False
        Workbooks.OpenXML varArquivo, , xlXmlLoadImportToList
        Application.DisplayAlerts = True
        If CStr(Cells(1, 9)) = "ns1:numeroDemonstrativo" Then
            varDemonstrativo = CStr(Cells(2, 9))
            varTransacao = CStr(Cells(2, 2))
        Else
            varDemonstrativo = CStr(Cells(1, 9))
            varTransacao = CStr(Cells(1, 2))
        End If
        ActiveWorkbook.Close (False)
    
        indW = 1
        While Sheets("Download").Cells(indW, 2) <> ""

            If Sheets("Download").Cells(indW, 2) = varDemonstrativo Then
        
                varData = Format(Sheets("Download").Cells(indW, 3), "yyyymmdd")
                varArquivoNovo = pPasta & "DAC_" & varData & "_" & varTransacao & ".xml"
                Debug.Print varArquivoNovo
                Name varArquivo As varArquivoNovo
                GoTo Saida_DAC
            End If
            indW = indW + 1
    
        Wend
    
Saida_DAC:
        varArquivosDir = Dir()
    
    Loop
    '
    '------------------------------------------------------------------------------------------------
    '   Renomeia PAG
    '
    varArquivosDir = Dir(pPasta & "demonstrativoPgtoXml_*.xml")

    Do While varArquivosDir <> ""
        varArquivo = pPasta & CStr(varArquivosDir)
    
        Application.DisplayAlerts = False
        Workbooks.OpenXML varArquivo, , xlXmlLoadImportToList
        Application.DisplayAlerts = True
        If CStr(Cells(1, 9)) = "ns1:numeroDemonstrativo" Then
            varDemonstrativo = CStr(Cells(2, 9))
            varData = Format(CStr(Cells(2, 16)), "yyyymmdd")
        Else
            varDemonstrativo = CStr(Cells(1, 9))
            varData = Format(CStr(Cells(1, 16)), "yyyymmdd")
        End If
        ActiveWorkbook.Close (False)
    
        varArquivoNovo = pPasta & "PAG_" & varData & "_" & varDemonstrativo & ".xml"
        Debug.Print varArquivoNovo
    
        Name varArquivo As varArquivoNovo

        varArquivosDir = Dir()
    Loop

End Sub

Sub teste()

    Unzip "C:\Users\renat\Downloads\health TEC\digest\CASSE\"
    RenomearArquivos "C:\Users\renat\Downloads\health TEC\digest\CASSE\"
End Sub

