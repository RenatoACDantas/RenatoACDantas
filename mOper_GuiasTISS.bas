Attribute VB_Name = "mOper_GuiasTISS"
'@Folder "Operadoras"
Option Explicit

Public Sub GuiasTISS(pLogin As String, pPass As String)
    On Error GoTo GuiasTISS_ERR
 
    ' Declarações

    Dim por As New By
    
    Dim ind As Integer
    Dim mChData As Integer

    Set driver = New ChromeDriver

redo:
    
    '
    '------------------------------------------------------------------------------------------------------
    '   Propriedades do browser
    '
    '    driver.setPreference "download.default_directory", mPasta
    driver.setPreference "download.directory_upgrade", True
    driver.setPreference "download.prompt_for_download", False
    



    
    '    driver.setProfile "C:\Users\renat\AppData\Local\Google\Chrome\User Data"
    '

    '------------------------------------------------------------------------------------------------------
    '   Navegar
    '
    If pLogin = "" Or pPass = "" Then GoTo redo
    driver.Start "chrome"

    driver.Window.Maximize
     
    driver.Get "https://www.guiastiss.com.br/"
    driver.Wait 500
    
    driver.FindElementByCss("#u2259-4 > p:nth-child(1)").Click
    driver.Wait 100
    
    driver.FindElementByCss("#MainContent_lgAutenticacao_UserName").Clear
    driver.FindElementById("MainContent_lgAutenticacao_UserName").SendKeys pLogin

    driver.FindElementById("MainContent_lgAutenticacao_Password").Clear
    driver.FindElementById("MainContent_lgAutenticacao_Password").SendKeys pPass
    
    driver.FindElementByCss("#MainContent_lgAutenticacao_btnValidarLogin").Click
    driver.Wait 500
        
    While Not driver.isElementPresent(por.Css("#lnkRelatorios > a"))
        Debug.Print Now
        Application.Wait Now + TimeValue("00:00:01")
        driver.Wait 1000
    Wend

    driver.FindElementByCss("#lnkRelatorios > a").ReleaseMouse
    driver.Wait 100
    
    driver.FindElementByCss("#lnkRelatorios > ul > li:nth-child(1) > a").Click
    driver.Wait 500
    
    driver.FindElementByCss("#MainContent_rblOpcao_1").Click
    driver.Wait 500
    '
    '----------------------------------------------------------------------------------
    '   Ler parâmetros de datas
    '

    ind = 2
    While Sheets("Parametros").Cells(ind, 1) <> ""
   
        mChData = Sheets("Parametros").Cells(ind, 1)
        mDataInicial = Replace(Sheets("Parametros").Cells(ind, 2), "/", "")
        mDataFinal = Replace(Sheets("Parametros").Cells(ind, 3), "/", "")

        driver.FindElementById("MainContent_tbDateFrom").SendKeys mDataInicial
        driver.Wait 100

        driver.FindElementById("MainContent_tbDateTo").SendKeys mDataFinal
        driver.Wait 1000
    
        driver.FindElementById("MainContent_lnkRelatorio").Click
        driver.Wait 9000

        '
        '-------------------------------------------------------------------------------------------
        '  Navegar para baixar Demonstrativos
        '
    
        driver.SwitchToNextWindow.Activate
        '
        '--------------------------------------------------------------------------------------
        '   Baixo arquivo em PDF
    
        driver.FindElementByCss("#reportViewer_Splitter_Toolbar_Menu_DXI9_Img").Click
        driver.Wait 900

        '
        '--------------------------------------------------------------------------------------------------------
        '   Alterar para XLSX e baixar arquivo
        '
        driver.FindElementByCss("#reportViewer_Splitter_Toolbar_Menu_ITCNT11_SaveFormat_I").Click
        driver.Wait 1000
   
        driver.FindElementByCss("#reportViewer_Splitter_Toolbar_Menu_ITCNT11_SaveFormat_DDD_L_LBI2T0").Click
        driver.Wait 1000
        
        driver.FindElementByCss("#reportViewer_Splitter_Toolbar_Menu_DXI9_Img").Click
        driver.Wait 500
        
        driver.SwitchToPreviousWindow.Activate
        
        ind = ind + 1
      
    Wend
    '
    '--------------------------------------------------------------------------------------------------
    '   Fechar browse
    '
    driver.Quit
    Sheets("Menu").Select
    Exit Sub
    '
GuiasTISS_ERR:
    MsgBox Error$
    Sheets("Menu").Select
    driver.Quit

End Sub

