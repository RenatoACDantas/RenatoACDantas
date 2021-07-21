Attribute VB_Name = "mImportar_arquivos"
'@Folder "Importa��o"
Option Explicit

Sub Importar_Arquivo(Se��o As String)
    On Error GoTo Importar_Arquivo_Err


    Dim fDialog As Office.FileDialog
    Dim varFile As Variant
    Dim varArquivoDir As Variant
    Dim varArquivosDir As Variant
    Dim mExtens�o As String
    Dim mCr�ticaArquivo As String
    Dim mCr�ticaPosi��o As String
    Dim mCr�ticaPosi��o_2 As String
        
    Dim ind As Integer
 
    Application.ScreenUpdating = False
    '    Application.Cursor = xlWait

    Montar_listbox "In�cio", Se��o
      
    mArquivoAberto = ActiveWorkbook.Name
    
    mPastaPadrao = Pesquisar_chave("Geral", "Pasta Padr�o")
    
    mCr�ticaArquivo = Pesquisar_chave(Se��o, "Cr�ticaArquivo")
    mCr�ticaPosi��o = Pesquisar_chave(Se��o, "Cr�ticaPosi��o")
    mCr�ticaPosi��o_2 = Pesquisar_chave(Se��o, "Cr�ticaPosi��o_2")
    mExtens�o = Pesquisar_chave(Se��o, "Extens�o")
    
    '    Set fDialog = Application.FileDialog(msoFileDialogOpen)
    '    With fDialog
    '      .InitialFileName = mPasta
    '      .Title = "Escolha o arquivo para importar"
    '      .Filters.Clear
    '      .Filters.Add Se��o, mExtens�o
    '      If .Show = True Then
    '         For Each varFile In .SelectedItems

    Debug.Print Replace(mPasta & Se��o & mExtens�o, mSe��oInterna & "-", "")
    varArquivosDir = Dir(Replace(mPasta & Se��o & mExtens�o, mSe��oInterna & "-", ""))

    ind = 0
    Do While varArquivosDir <> ""
        mArquivo = mPasta & CStr(varArquivosDir)
        
        Debug.Print mArquivo
        '
        '------------------------------------------------------------------------------------
        '   Validar arquivo
        '
        If mExtens�o = "*.XML" Then
            Application.DisplayAlerts = False
            Workbooks.OpenXML mArquivo, , xlXmlLoadImportToList
            Application.DisplayAlerts = True
        Else
            Workbooks.Open mArquivo
        End If
        Sheets(1).Select
                
        If Range(mCr�ticaPosi��o) = mCr�ticaArquivo Or _
                                    Range(mCr�ticaPosi��o_2) = mCr�ticaArquivo Then

            ActiveWorkbook.Close (False)
            Sheets("Menu").Select
            '
            mReferencia = Replace(mArquivo, mPasta, "")
            mReferencia = Replace(Mid(mReferencia, 1, InStr(1, mReferencia, ".") - 1), " ", "")
            UserForm_Zscan.tbReferencia = mReferencia
            UserForm_Zscan.tbTotLaudos_O = ""
            UserForm_Zscan.tbTotLaudos_E = ""

            '                Gravar_chave Se��o, "Pasta", fDialog.InitialFileName
            Gravar_chave Se��o, "Arquivo", mArquivo
            Gravar_chave mTela, "Refer�ncia", mReferencia
            Gravar_chave Se��o, "ArquivoAberto", mArquivoAberto
            
            mArquivoImport = mArquivoBase
            mNomeModelo = Pesquisar_chave(Se��o, "NomeModelo")
                
            mPasta = Pesquisar_chave(Se��o, "Pasta") & mPastaOperadora
            mArquivoSaidaName = Replace(mNomeModelo, "XXX", mReferencia)
            mArquivoSaida = mPasta + mArquivoSaidaName
            mArquivoBase = Pesquisar_chave(Se��o, "ArquivoBase")
            mPlanilhaTotal = Pesquisar_chave(Se��o, "PlanilhaTotal")
            mPlanilhaTotalEditado = Pesquisar_chave(Se��o, "PlanilhaTotalEditado")
            mArquivoModelo = Pesquisar_chave(Se��o, "Modelo")
            mReferencia = Pesquisar_chave(Se��o, "Refer�ncia")
            mRange = Pesquisar_chave(Se��o, "Range")


            Gravar_chave Se��o, "ArquivoSaidaName", mArquivoSaidaName
            Gravar_chave Se��o, "ArquivoSaida", mArquivoSaida
            
            Montar_listbox "Pasta:", mPasta
            Montar_listbox "Planilha:", mArquivo
            
            Montar_listbox "Mensagem: ", "Planilha importada com sucesso"
            
            Sheets("Menu").Select
            '
            '---------------------------------------------------------------------------------------
            '   Criar e editar planilha
            '
            '               Application.Cursor = xlWait
            Criar_Arquivo_Editado (Se��o)
            Application.Cursor = xlDefault
            '
            '---------------------------------------------------------------------------------------
            '   Incluir dados em banco de dados
            '
            '               Application.Cursor = xlWait
            Incluir_BD (Se��o)
            Application.Cursor = xlDefault
            '
            '---------------------------------------------------------------------------------------
            '   Arquivo inv�lido
            '
        Else
            MsgBox Range(mCr�ticaPosi��o) + " - " + mCr�ticaArquivo + " - " + Range(mCr�ticaPosi��o_2)
            ActiveWorkbook.Close (False)
            Windows(mArquivoAberto).Activate
            Sheets("Menu").Select
            Montar_listbox "Erro", mArquivo + " - Arquivo selecionado inv�lido"
            MsgBox "Arquivo selecionado inv�lido", vbCritical
            

        End If

        '        Next
        '
        Beep
        Montar_listbox "T�rmino: ", "Planilhas processadas com sucesso"
        UserForm_Zscan.tbTotLaudos_O = ""
        UserForm_Zscan.tbTotLaudos_E = ""


        '    Else
        '        Beep
        '       Montar_listbox "Aviso: ", "Nenhum arquivo selecionado"
        '       MsgBox "Nenhum arquivo selecionado", vbExclamation
      
        '      End If
      
        ind = ind + 1
        varArquivosDir = Dir()
    Loop
   
    Montar_listbox "T�rmimo importa��o: ", mTela
    Montar_listbox "-----", "-----"

Importar_Arquivo_Exit:
    Exit Sub

Importar_Arquivo_Err:

    Montar_listbox "Erro importa: ", Err.Description

    MsgBox Error$
    ActiveWorkbook.Save
    
    '  Resume Importar_Arquivo_Exit

End Sub

Public Sub Montar_listbox(coluna1 As String, coluna2 As String)
    Debug.Print coluna1 & " " & coluna2 & " " & mTela
    With UserForm_Zscan.MultiPage1.Pages("pgLog").ListBoxLog
        .AddItem
        .List(mInd, 0) = coluna1
        .List(mInd, 1) = coluna2
    End With

    Sheets("Log").Cells(mIndLog, 1) = coluna1
    Sheets("Log").Cells(mIndLog, 2) = coluna2

    mIndLog = mIndLog + 1
    mInd = mInd + 1
End Sub

Sub t()
TaskManager "chrome.exe"
End Sub

