Attribute VB_Name = "mMenuRibbom"
'@Folder "Ambiente"
Option Explicit

'Callback for customButton1 onAction
Sub conImportZscanSub(control As IRibbonControl)
    mTela = "Zscan"
    UserForm_Zscan.Show
End Sub

'Callback for customButton2 onAction
Sub conImportGuiasTissSub(control As IRibbonControl)
    mTela = "Guias"
    UserForm_Zscan.Show
End Sub

'Callback for customButton3 onAction
Sub conImportOperadorasSub(control As IRibbonControl)
    mTela = "Operadora"
    UserForm_Zscan.Show
End Sub

'Callback for customButton11 onAction
Sub conConcilaçãoSub(control As IRibbonControl)
    mPastaPadrao = Pesquisar_chave("Geral", "Pasta Padrão")
    mArquivo = mPastaPadrao & "Resumo.xlsb"
    Workbooks.Open mArquivo
    Sheets("Conciliação").Select
End Sub

'Callback for customButton12 onAction
Sub conPagamentosSub(control As IRibbonControl)
    mPastaPadrao = Pesquisar_chave("Geral", "Pasta Padrão")
    mArquivo = mPastaPadrao & "Resumo.xlsb"
    Debug.Print mArquivo
    Workbooks.Open mArquivo
    Sheets("Pagamentos").Select
End Sub

'Callback for customButton13 onAction
Sub conConexãoSub(control As IRibbonControl)
    Sheets("conexao").Select
End Sub

'Callback for customButton21 onAction
Sub conConfiguraçãoSub(control As IRibbonControl)
    Sheets("Config_2").Select
End Sub

'Callback for customButton22 onAction
Sub conAcessosSub(control As IRibbonControl)
    Sheets("Acessos").Select
End Sub

'Callback for customButton24 onAction
Sub conParâmetroSub(control As IRibbonControl)
    Sheets("Parametros").Select
End Sub

'Callback for customButton31 onAction
Sub conArquivoSub(control As IRibbonControl)
    Sheets("Arquivos").Select
End Sub

'Callback for customButton32 onAction
Sub conLogSub(control As IRibbonControl)
    Sheets("Log").Select
End Sub

'Callback for customButton33 onAction
Sub conDownloadSub(control As IRibbonControl)
    Sheets("Download").Select
End Sub

