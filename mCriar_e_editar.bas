Attribute VB_Name = "mCriar_e_editar"
'@Folder "Importa��o"
Option Explicit

Sub Criar_Arquivo_Editado(Se��o As String)
    On Error GoTo Criar_Arquivo_editado_Err

    ' Declare��es

    Dim varColuna As String
    Dim varRange As String
    
    Dim mIndRow As Integer
    Dim mIndConexoes As Integer
    Dim mIndQueries As Integer
    Dim mQueries As Integer
    

    Montar_listbox " ", " "
    Application.ScreenUpdating = False
    
    Sheets("Menu").Select

    Montar_listbox "Planilha editada:", mArquivoSaidaName

    FileCopy mPastaPadrao & mArquivoModelo, mArquivoSaida
    Workbooks.Open mArquivoSaida
    '
    '------------------------------------------------------------------------------------
    '   AtualizaConex�o

    Configurar_Conexao Se��o
    ActiveWorkbook.RefreshAll
    Montar_listbox "Mensagem:", "Planilhas atualizadas"

    ActiveWorkbook.Save
     
    '
    '--------------------------------------------------------------------------------------
    '   Incluir coluna de Refer�ncia
    '

    Windows(mArquivoSaidaName).Activate
    Sheets(mArquivoBase).Select
    

    Debug.Print "Range: " & mTela & " " & mRange & " " & InStr(1, mRange, ":", vbTextCompare) & " " & Mid(mRange, InStr(1, mRange, ":", vbTextCompare) + 1, 1)
    '    If mTela = "Zscan" Then
    '        varColuna = "I"
    '    ElseIf mTela = "Guias" Then
    '        varColuna = "J"
    '    ElseIf mTela = "XML-DAC" Then
    '        varColuna = "O"
    '    ElseIf mTela = "XML-PAG" Then
    '        varColuna = "P"
    '    ElseIf mTela = "CASSE" Then
    '        varColuna = "M"
    '    End If

    varColuna = Mid(mRange, InStr(1, mRange, ":", vbTextCompare) + 1, 1)

    Range(varColuna + "3") = "Refer�ncia"
    Range(varColuna + "4") = mReferencia
    If Range("A5") <> "" Then
        Range(varColuna + "5") = mReferencia
    End If
    Range("A3").Select                           '--
    Range(Selection, Selection.End(xlDown)).Select
    mIndRow = Selection.End(xlDown).Row

    varRange = varColuna + "4:" + varColuna + CStr(mIndRow)
    If mIndRow > 5 Then
        Range(varColuna + "4:" + varColuna + "5").Select
        Selection.AutoFill Destination:=Range(varRange)
    End If

    Windows(mArquivoAberto).Activate
    varRange = "A3:" + varColuna + CStr(mIndRow)
    Gravar_chave Se��o, "Range", varRange
 
    '
    '------------------------------------------------------------------------------------
    '   Atualizar tabela din�mico Total Editado
    '
    Windows(mArquivoSaidaName).Activate
    Sheets(mPlanilhaTotalEditado).Select
    ActiveSheet.PivotTables("Tabela din�mica1").PivotCache.Refresh
    '
    '--------------------------------------------------------------------------------------
    '   Montar totais editados
    '
    If mTela = "Zscan" Then
        mTotalEditado = CStr(Cells(2, 1))
    ElseIf mTela = "Guias" Then
        mTotalEditado = CStr(FormatCurrency(Cells(2, 1)))
    ElseIf mTela = "XML-DAC" Then
        mTotalEditado = CStr(FormatCurrency(Cells(3, 4)))
    ElseIf mTela = "XML-PAG" Then
        mTotalEditado = CStr(FormatCurrency(Cells(3, 4)))
    ElseIf mTela = "CASSE" Then
        mTotalEditado = CStr(FormatCurrency(Cells(2, 3)))
    Else
        mTotalEditado = CStr(FormatCurrency(Cells(2, 4)))
    End If
    '
    '--------------------------------------------------------------------------------------
    '   Montar totais originais
    '
    Sheets(mPlanilhaTotal).Select
            
    If Se��o = "Zscan" Then
        If IsNumeric(Cells(2, 2)) Then
            mTotal = Cells(2, 2)
        ElseIf IsNumeric(Cells(2, 3)) Then
            mTotal = Cells(2, 3)
        End If
    ElseIf Se��o = "Guias" Then
        mTotal = FormatCurrency(Cells(2, 1))
    ElseIf Se��o = "XML-DAC" Then
        mTotal = FormatCurrency(Cells(2, 12))
    ElseIf Se��o = "XML-PAG" Then
        mTotal = FormatCurrency(Cells(2, 10))
    ElseIf Se��o = "UNIMED-PAG" Then
        mTotal = FormatCurrency(Cells(2, 4))
    ElseIf Se��o = "CASSE" Then
        mTotal = FormatCurrency(Cells(2, 3))
    Else
        mTotal = FormatCurrency(Cells(2, 6))
    End If
    '
    '-----------------------------------------------------------------------------------------
    '   Desabilitar conex�es dos novos arquivos criados e editados
    '
    mConexoes = ActiveWorkbook.Connections.Count
 
    For mIndConexoes = 1 To mConexoes
        '       ActiveWorkbook.Connections(mIndConexoes).Delete
    Next mIndConexoes
    
    mQueries = ActiveWorkbook.Queries.Count
 
    For mIndQueries = 1 To mQueries
        '        ActiveWorkbook.Queries(mIndQueries).Delete
    Next mIndQueries
    
    Range("A1").Select
    
    ' ActiveWorkbook.Save

    '--------------------------------------------------------------------------------------
    Sheets(mArquivoBase).Select
  
    ActiveWorkbook.Save
    ActiveWorkbook.Close
    '
    '--------------------------------------------------------------------------------------
    '   Exibir totais
    '
    Windows(mArquivoAberto).Activate
    Sheets("Menu").Select
    Application.Cursor = xlDefault
    Montar_listbox "Total de Laudos:", mTotal
    UserForm_Zscan.tbTotLaudos_O.Value = mTotal
    
    Montar_listbox "Total de Laudos:", mTotalEditado
    UserForm_Zscan.tbTotLaudos_E.Value = mTotalEditado
     
    Montar_listbox "Mensagem:", "Planilha editada gerada com sucesso"
    '
    '---------------------------------------------------------------------------------------
    '   Montar hiperlink do arquivo criado
    '
    Sheets("Arquivos").Select
    
    Sheets("Arquivos").Cells(mIndPlanilha, 1) = Now
    Sheets("Arquivos").Cells(mIndPlanilha, 2) = Se��o
    Sheets("Arquivos").Cells(mIndPlanilha, 3) = mOperadora

    Sheets("Arquivos").Cells(mIndPlanilha, 4) = mArquivo
    Sheets("Arquivos").Cells(mIndPlanilha, 4).Select

    ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:= _
                               Replace(mArquivo, " ", " "), TextToDisplay:=Replace(mArquivo, mPasta, "")
       
    Sheets("Arquivos").Cells(mIndPlanilha, 5) = mArquivoSaidaName
    Sheets("Arquivos").Cells(mIndPlanilha, 5).Select

    ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:= _
                               Replace(mPasta & mArquivoSaidaName, " ", " "), TextToDisplay:=mArquivoSaidaName
       
    Sheets("Arquivos").Cells(mIndPlanilha, 6) = mTotalEditado
       
    mIndPlanilha = mIndPlanilha + 1
    Montar_listbox "Mensagem:", "Hiperlink criado"
    Sheets("Menu").Select

    '   UserForm_Zscan.WebBrowser1.Navigate mArquivoSaida
    
Criar_Arquivo_editado_Exit:
    Montar_listbox "T�rmino cria��o e edi��o", Se��o
    Montar_listbox "-----", "-----"
    Exit Sub

Criar_Arquivo_editado_Err:
    Montar_listbox "Erro: ", Err.Description

    MsgBox Error$

    ActiveWorkbook.Save
    ' Resume Criar_Arquivo_editado_Exit
            
End Sub


