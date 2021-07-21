Attribute VB_Name = "mImportar_arquivos"
'@Folder "Importação"
Option Explicit

Sub Importar_Arquivo(Seção As String)
    On Error GoTo Importar_Arquivo_Err


    Dim fDialog As Office.FileDialog
    Dim varFile As Variant
    Dim varArquivoDir As Variant
    Dim varArquivosDir As Variant
    Dim mExtensão As String
    Dim mCríticaArquivo As String
    Dim mCríticaPosição As String
    Dim mCríticaPosição_2 As String
        
    Dim ind As Integer
 
    Application.ScreenUpdating = False
    '    Application.Cursor = xlWait

    Montar_listbox "Início", Seção
      
    mArquivoAberto = ActiveWorkbook.Name
    
    mPastaPadrao = Pesquisar_chave("Geral", "Pasta Padrão")
    
    mCríticaArquivo = Pesquisar_chave(Seção, "CríticaArquivo")
    mCríticaPosição = Pesquisar_chave(Seção, "CríticaPosição")
    mCríticaPosição_2 = Pesquisar_chave(Seção, "CríticaPosição_2")
    mExtensão = Pesquisar_chave(Seção, "Extensão")
    
    '    Set fDialog = Application.FileDialog(msoFileDialogOpen)
    '    With fDialog
    '      .InitialFileName = mPasta
    '      .Title = "Escolha o arquivo para importar"
    '      .Filters.Clear
    '      .Filters.Add Seção, mExtensão
    '      If .Show = True Then
    '         For Each varFile In .SelectedItems

    Debug.Print Replace(mPasta & Seção & mExtensão, mSeçãoInterna & "-", "")
    varArquivosDir = Dir(Replace(mPasta & Seção & mExtensão, mSeçãoInterna & "-", ""))

    ind = 0
    Do While varArquivosDir <> ""
        mArquivo = mPasta & CStr(varArquivosDir)
        
        Debug.Print mArquivo
        '
        '------------------------------------------------------------------------------------
        '   Validar arquivo
        '
        If mExtensão = "*.XML" Then
            Application.DisplayAlerts = False
            Workbooks.OpenXML mArquivo, , xlXmlLoadImportToList
            Application.DisplayAlerts = True
        Else
            Workbooks.Open mArquivo
        End If
        Sheets(1).Select
                
        If Range(mCríticaPosição) = mCríticaArquivo Or _
                                    Range(mCríticaPosição_2) = mCríticaArquivo Then

            ActiveWorkbook.Close (False)
            Sheets("Menu").Select
            '
            mReferencia = Replace(mArquivo, mPasta, "")
            mReferencia = Replace(Mid(mReferencia, 1, InStr(1, mReferencia, ".") - 1), " ", "")
            UserForm_Zscan.tbReferencia = mReferencia
            UserForm_Zscan.tbTotLaudos_O = ""
            UserForm_Zscan.tbTotLaudos_E = ""

            '                Gravar_chave Seção, "Pasta", fDialog.InitialFileName
            Gravar_chave Seção, "Arquivo", mArquivo
            Gravar_chave mTela, "Referência", mReferencia
            Gravar_chave Seção, "ArquivoAberto", mArquivoAberto
            
            mArquivoImport = mArquivoBase
            mNomeModelo = Pesquisar_chave(Seção, "NomeModelo")
                
            mPasta = Pesquisar_chave(Seção, "Pasta") & mPastaOperadora
            mArquivoSaidaName = Replace(mNomeModelo, "XXX", mReferencia)
            mArquivoSaida = mPasta + mArquivoSaidaName
            mArquivoBase = Pesquisar_chave(Seção, "ArquivoBase")
            mPlanilhaTotal = Pesquisar_chave(Seção, "PlanilhaTotal")
            mPlanilhaTotalEditado = Pesquisar_chave(Seção, "PlanilhaTotalEditado")
            mArquivoModelo = Pesquisar_chave(Seção, "Modelo")
            mReferencia = Pesquisar_chave(Seção, "Referência")
            mRange = Pesquisar_chave(Seção, "Range")


            Gravar_chave Seção, "ArquivoSaidaName", mArquivoSaidaName
            Gravar_chave Seção, "ArquivoSaida", mArquivoSaida
            
            Montar_listbox "Pasta:", mPasta
            Montar_listbox "Planilha:", mArquivo
            
            Montar_listbox "Mensagem: ", "Planilha importada com sucesso"
            
            Sheets("Menu").Select
            '
            '---------------------------------------------------------------------------------------
            '   Criar e editar planilha
            '
            '               Application.Cursor = xlWait
            Criar_Arquivo_Editado (Seção)
            Application.Cursor = xlDefault
            '
            '---------------------------------------------------------------------------------------
            '   Incluir dados em banco de dados
            '
            '               Application.Cursor = xlWait
            Incluir_BD (Seção)
            Application.Cursor = xlDefault
            '
            '---------------------------------------------------------------------------------------
            '   Arquivo inválido
            '
        Else
            MsgBox Range(mCríticaPosição) + " - " + mCríticaArquivo + " - " + Range(mCríticaPosição_2)
            ActiveWorkbook.Close (False)
            Windows(mArquivoAberto).Activate
            Sheets("Menu").Select
            Montar_listbox "Erro", mArquivo + " - Arquivo selecionado inválido"
            MsgBox "Arquivo selecionado inválido", vbCritical
            

        End If

        '        Next
        '
        Beep
        Montar_listbox "Término: ", "Planilhas processadas com sucesso"
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
   
    Montar_listbox "Térmimo importação: ", mTela
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

