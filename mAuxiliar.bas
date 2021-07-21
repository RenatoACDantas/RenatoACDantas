Attribute VB_Name = "mAuxiliar"
'@Folder "Importação"
Option Explicit

Public mArquivoSaida As String
Public mResumo As Boolean
Public mPaint As Boolean
Public mArquivo As String
Public mArquivoImport As String
Public mArquivoModelo As String
Public mArquivoAberto As String
Public mArquivoSaidaName As String
Public mLogo As String
Public mInd As Integer
Public mIndLog As Integer
Public mIndPlanilha As Integer
Public mArquivoBase As String
Public mReferencia As String
Public mNomeModelo As String
Public mPastaPadrao As String
Public mPasta As String
Public mPastaOperadora As String
Public mOperadora As String
Public mSeçãoInterna As String

Public mArquivoEditado As Object
Public mTela As String
Public mPlanilhaTotal As String
Public mPlanilhaTotalEditado As String
Public mCriticaArquivo As String
Public mCriticaPosição As String
Public mCriticaPosição_2 As String
Public mTabela As String
Public mRange As String
Public mTotalEditado As String
Public mTotal As String

Public mConexoes

Sub Inicializar_Variaveis()

    ' Declarações

    Dim ind As Integer
    Dim mIndLog As Integer
    Dim mRows As String

    'Sheets("Log").Select
    'Range("A1").Select
    'Range(Selection, Selection.End(xlDown)).Select
    ''Selection.EntireRow.Delete

    mIndLog = 1

    Sheets("Arquivos").Select
    For ind = 18 To 25
        Cells(ind, 1) = ""
    Next ind
    Sheets("Menu").Select

End Sub

Sub Ler_Configuração()
     
    mLogo = Pesquisar_chave("Geral", "Logomarca")

    
End Sub

Sub Configurar_Conexao(Seção As String)
    '
    Dim mIndConexao As Integer
    Dim mQtConexoes As Integer
    Dim mSecao As String
    Dim mFormula As String
    Dim mConexao As String
    Dim mQuery As String

    Windows(mArquivoAberto).Activate
    Sheets("Conexao").Select
    Columns("A:A").Select
    Selection.Find(What:=Seção, After:=ActiveCell, LookIn:=xlFormulas2, _
                   LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                   MatchCase:=False, SearchFormat:=False).Activate
        
    mQtConexoes = ActiveCell.Row

    mIndConexao = mQtConexoes
    Debug.Print "Conexao: " & CStr(mIndConexao)
    While Cells(mIndConexao, 1) = Seção

        mFormula = Replace(Sheets("Conexao").Cells(mIndConexao, 2), "#mArquivoBase", mArquivo)
        mQuery = Sheets("Conexao").Cells(mIndConexao, 3)
        Debug.Print mFormula
        Debug.Print mQuery
    
        Windows(mArquivoSaidaName).Activate
        ActiveWorkbook.Queries(mQuery).Formula = mFormula

        mIndConexao = mIndConexao + 1
        Windows(mArquivoAberto).Activate
        Sheets("Conexao").Select
    Wend
    Montar_listbox "Mensagem: ", "Consultas atualizadas"
    mIndConexao = mQtConexoes
    Debug.Print mArquivoAberto & " " & mArquivoSaidaName

    Windows(mArquivoAberto).Activate
    Sheets("Conexao").Select
    While Cells(mIndConexao, 1) = Seção

        mConexao = Sheets("Conexao").Cells(mIndConexao, 4)

        Windows(mArquivoSaidaName).Activate
        ActiveWorkbook.Connections(mConexao).Refresh
        mIndConexao = mIndConexao + 1
    
        Windows(mArquivoAberto).Activate
        Sheets("Conexao").Select
    Wend
    Windows(mArquivoSaidaName).Activate
    ActiveWorkbook.RefreshAll
    Windows(mArquivoAberto).Activate

    Montar_listbox "Mensagem: ", "Conexões atualizadas"
End Sub

Sub Montar_cabeçalho(nome As String)

    ActiveWindow.DisplayGridlines = False
    ActiveSheet.Name = nome
    
    Columns("C:C").Select
    Range("C3").Activate
    Selection.ColumnWidth = 3
    Columns("F:F").Select
    Range("F3").Activate
    Selection.ColumnWidth = 3
    
    ActiveSheet.PageSetup.LeftHeaderPicture.Filename = _
                                                     mLogo
    With ActiveSheet.PageSetup.LeftHeaderPicture
        .Height = 42.75
        .Width = 98.25
    End With
    With ActiveSheet.PageSetup
        .Orientation = xlLandscape
        .LeftHeader = "&G"
        .CenterHeader = "&""-,Negrito""&14Resumos"
        .RightHeader = "&P/&N"
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = "Gerada em: &D &T"
    End With
End Sub

Sub Gravar_chave(pSeção As String, pChave As String, pValor As String)
    On Error GoTo Gravar_chave_Err

    Dim mSheetAnterior As String
    Dim mLinha As Integer
    
    mSheetAnterior = ActiveSheet.Name

    Sheets("Config_2").Select

    mLinha = 6
    For mLinha = 6 To 999


        If Cells(mLinha, 2) = pSeção And _
                              Cells(mLinha, 3) = pChave Then
            GoTo Gravar_chave_Exit
        
        ElseIf Cells(mLinha, 2) = "" Then
            GoTo Gravar_chave_Inc
        End If

    Next mLinha

Gravar_chave_Inc:
    Cells(mLinha, 2) = pSeção
    Cells(mLinha, 3) = pChave
    Cells(mLinha, 4) = pValor
    
    With ActiveSheet.ListObjects("Tabela134").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ActiveSheet.ListObjects("Tabela134").Range.AutoFilter Field:=4, Criteria1:="<>"
    
    Sheets(mSheetAnterior).Select
    
    Exit Sub
    
Gravar_chave_Exit:
    Cells(mLinha, 4) = pValor
    Sheets(mSheetAnterior).Select

    Exit Sub

Gravar_chave_Err:
    MsgBox "G - Chave inválida - " + pSeção + " - " + pChave + " - " + pValor, vbCritical
    Sheets(mSheetAnterior).Select
    
End Sub

Sub Montar_listbox_do_log()
    On Error GoTo Montar_listbox_do_log_Err

    Sheets("Log").Select
    Range("A1").Select
    mIndLog = Selection.End(xlDown).Row
   
    If mIndLog > 1 Then
        For mInd = 0 To mIndLog - 1
            With UserForm_Zscan.MultiPage1.Pages("pgLog").ListBoxLog
                .AddItem
                .List(mInd, 0) = Sheets("Log").Cells(mInd + 1, 1)
                .List(mInd, 1) = Sheets("Log").Cells(mInd + 1, 2)
            End With
        Next mInd
    End If

Montar_listbox_do_log_Exit:

    Sheets("Menu").Select
    Exit Sub

Montar_listbox_do_log_Err:

    Resume Montar_listbox_do_log_Exit

End Sub

Function Pesquisar_chave(pSeção As String, pChave As String) As String
    On Error GoTo Pesquisar_chave_Err

    Dim mLinha As Integer
    Dim mValor As String
    Dim mSheetAnterior As String

    mSheetAnterior = ActiveSheet.Name

    Sheets("Config_2").Select

    mLinha = 6
    For mLinha = 6 To 999

        If Cells(mLinha, 2) = pSeção And _
                              Cells(mLinha, 3) = pChave Then
            GoTo Pesquisar_chave_Exit
        
        ElseIf Cells(mLinha, 2) = "" Then

            GoTo Pesquisar_chave_Err
        End If

    Next mLinha

Pesquisar_chave_Exit:
    mValor = Cells(mLinha, 4)
    Sheets(mSheetAnterior).Select
    
    Pesquisar_chave = mValor

    Exit Function

Pesquisar_chave_Err:
    MsgBox "P - Chave inválida - " + pSeção + " - " + pChave, vbCritical
    Sheets(mSheetAnterior).Select

End Function

Function Alterar_PastaPadrão(pPastaPadrão As String)

    Dim varPasta As String
    Dim indW As Integer
    Debug.Print "Alterar: " & pPastaPadrão
    indW = 6
    While Sheets("Config_2").Cells(indW, 2) <> ""
        Debug.Print CStr(indW)
        If Sheets("Config_2").Cells(indW, 2) <> "Geral" And _
                                             Sheets("Config_2").Cells(indW, 3) = "Pasta" Then
        
            Sheets("Config_2").Cells(indW, 4) = pPastaPadrão
            Debug.Print "Alterado: " & Sheets("Config_2").Cells(indW, 4)

        End If

        indW = indW + 1

    Wend

End Function

Function PastaOperadora(pSeção As String, pOperadora As String)
    Dim varSubPasta As String
    Dim varPasta As String
    Dim indW As Integer

    Debug.Print "Funct: " & pSeção & " " & pOperadora
    indW = 6
    While Sheets("Acessos").Cells(indW, 2) <> ""
  
        If Sheets("Acessos").Cells(indW, 2) = pOperadora Then
            varSubPasta = Sheets("Acessos").Cells(indW, 7)
            varPasta = Pesquisar_chave(pSeção, "Pasta") & varSubPasta
        
            PastaOperadora = varPasta
            Exit Function
        End If

        indW = indW + 1

    Wend

End Function

Public Function ArrayLen(arr As Variant) As Integer
    
    ArrayLen = UBound(arr) - LBound(arr) + 1
End Function

Public Function Crypt(Text As String) As String
    On Error GoTo Crypt_Err
    'Criptografia da senha

    Dim i

    Dim strTempChar As String

    For i = 1 To Len(Text)

        If Asc(Mid$(Text, i, 1)) < 128 Then
            strTempChar = Asc(Mid$(Text, i, 1)) + 128
        Else
            GoTo Crypt_Err
        End If

        Mid$(Text, i, 1) = Chr(strTempChar)

    Next i

    Crypt = Text

Crypt_Exit:
    Exit Function

Crypt_Err:
    Crypt = ""
    MsgBox "Texto não pode ser criptografado"

End Function

Public Function DeCrypt(Text As String) As String
    On Error GoTo DeCrypt_Err
    'Criptografia da senha

    Dim i

    Dim strTempChar As String

    For i = 1 To Len(Text)


        If Asc(Mid$(Text, i, 1)) > 128 Then
            strTempChar = Asc(Mid$(Text, i, 1)) - 128
        Else
            GoTo DeCrypt_Err
        End If

        Mid$(Text, i, 1) = Chr(strTempChar)

    Next i

    DeCrypt = Text

DeCrypt_Exit:
    Exit Function

DeCrypt_Err:
    DeCrypt = ""
    MsgBox "Texto não pode ser decriptografado"
    
End Function

Sub ExecutaBotões(pSeção As String, Optional pOperadora As String)
    On Error GoTo ExecutaBotões_Err
    
    ' Declarações

    Dim mMacro As String
    Dim mModulo As String
    Dim mLogin As String
    Dim mPass As String
    Dim mLoginAux As String

    Dim indW As Integer


    indW = 6
    While Sheets("Acessos").Cells(indW, 2) <> ""
        If Sheets("Acessos").Cells(indW, 2) = pSeção Or _
                                              Sheets("Acessos").Cells(indW, 2) = pOperadora Then
            mLogin = (Sheets("Acessos").Cells(indW, 3))
            mPass = DeCrypt(Sheets("Acessos").Cells(indW, 6))
            mMacro = Sheets("Acessos").Cells(indW, 5)
            mModulo = "mOper_" & mMacro
            mPastaOperadora = Sheets("Acessos").Cells(indW, 7)
            mLoginAux = Sheets("Acessos").Cells(indW, 9)
            mSeçãoInterna = Sheets("Acessos").Cells(indW, 8)

        
            mPasta = Pesquisar_chave(pSeção, "Pasta") & mPastaOperadora
        
            Debug.Print mPasta
            '
            '----------------------------------------------------------------------------------------------------
            '   Baixar arquivos
            '
            If UserForm_Zscan.OptionButton1.Value = True Or _
               UserForm_Zscan.OptionButton3.Value = True Then
               
                ' Verifica se google chrome está ativo. Deverá se encerrado
    
                If TaskManager("chrome.exe") = 2 Then 'Existe chrome ativo e não foi encerrado
                    Exit Sub
                End If
    
                ' Continua execução
    
                If mMacro = "UNIMED" Then
                    Application.Run "'" & Application.ThisWorkbook.FullName & "'!" & mModulo & "." & mMacro, mLogin, mLogin, mPass, pOperadora
                Else
                    Application.Run "'" & Application.ThisWorkbook.FullName & "'!" & mModulo & "." & mMacro, mLogin, mPass, pOperadora
                End If
                Sheets("Menu").Select


            End If

            '
            '-----------------------------------------------------------------------------------------------------
            '   Importar e editar arquivos
            '
            If UserForm_Zscan.OptionButton2.Value = True Or _
               UserForm_Zscan.OptionButton3.Value = True Then

                mIndPlanilha = 1
                If mTela = "Operadora" Then
                    mTela = mSeçãoInterna & "-DAC"
                    Importar_Arquivo mTela
                
                    mTela = mSeçãoInterna & "-PAG"
                    Importar_Arquivo mTela
                Else
                    Importar_Arquivo mTela
                End If
                Sheets("Menu").Select
            
            End If
        
            Exit Sub
        Else
            indW = indW + 1
        End If
    Wend
    
    MsgBox "Seção inválida: " & pSeção
ExecutaBotões_Fim:
    Exit Sub
    
ExecutaBotões_Err:
    Application.Cursor = xlDefault
    Montar_listbox "Google chrome: ", Err.Description

    MsgBox Error$
End Sub

Public Sub Montar_listbox_baixados(pOperadora As String)

    ' Delcarações

    Dim varPasta As String
    Dim varArquivo As String
    Dim varArquivosDir As Variant

    Dim ind As Integer

    varPasta = PastaOperadora(mTela, pOperadora)

    '
    '   Limpar lista
    '
    If UserForm_Zscan.MultiPage1.Pages("pgBaixados").ListBoxBaixados.ListCount > 1 Then
        For ind = 1 To (UserForm_Zscan.MultiPage1.Pages("pgBaixados").ListBoxBaixados.ListCount - 1) Step -1
            UserForm_Zscan.MultiPage1.Pages("pgBaixados").ListBoxBaixados.RemoveItem (ind)
        Next ind
    End If
    '
    '   Montar List
    '
    If mTela = "Operadora" Then
    
        '   DAC
        '
        If pOperadora = "UNIMED" Then
            varArquivosDir = Dir(varPasta & "DAC*.csv")
        Else
            varArquivosDir = Dir(varPasta & "DAC*.xml")
        End If
        ind = 0
        Do While varArquivosDir <> ""
            varArquivo = CStr(varArquivosDir)

            With UserForm_Zscan.MultiPage1.Pages("pgBaixados").ListBoxBaixados
                .AddItem
                .List(ind, 0) = varArquivo
            End With

            ind = ind + 1
    
            varArquivosDir = Dir()
        Loop
        '
        '   PAG
        '
        If pOperadora = "UNIMED" Then
            varArquivosDir = Dir(varPasta & "PAG*.xlsx")
        Else
            varArquivosDir = Dir(varPasta & "PAG*.xml")
        End If
        Do While varArquivosDir <> ""
            varArquivo = CStr(varArquivosDir)
            With UserForm_Zscan.MultiPage1.Pages("pgBaixados").ListBoxBaixados
                .AddItem
                .List(ind, 0) = varArquivo
            End With

            ind = ind + 1
    
            varArquivosDir = Dir()
        Loop
    Else
    
        '   Demais
        '
        varArquivosDir = Dir(varPasta & Pesquisar_chave(mTela, "ArquivoBase") & Pesquisar_chave(mTela, "Extensão"))

        ind = 0
        Do While varArquivosDir <> ""
            varArquivo = CStr(varArquivosDir)

            With UserForm_Zscan.MultiPage1.Pages("pgBaixados").ListBoxBaixados
                .AddItem
                .List(ind, 0) = varArquivo
            End With

            ind = ind + 1
    
            varArquivosDir = Dir()
        Loop
    
    End If
End Sub

Public Sub Montar_listbox_editados(pOperadora As String)

    ' Declarações

    Dim varPasta As String
    Dim varArquivo As String
    Dim varArquivosDir As Variant

    Dim ind As Integer

    varPasta = PastaOperadora(mTela, pOperadora)
    '
    '   Limpar lista
    '
    If UserForm_Zscan.MultiPage1.Pages("pgEditados").ListBoxEditados.ListCount > 1 Then
        For ind = 1 To (UserForm_Zscan.MultiPage1.Pages("pgEditados").ListBoxEditados.ListCount - 1) Step -1
            UserForm_Zscan.MultiPage1.Pages("pgEditados").ListBoxEditados.RemoveItem (ind)
        Next ind
    End If
    '
    '   Montar List
    '
    If mTela = "Operadora" Then
    
        '   DAC
        '
        varArquivosDir = Dir(varPasta & "ED_DAC*.xml")

        ind = 0
        Do While varArquivosDir <> ""
            varArquivo = CStr(varArquivosDir)

            With UserForm_Zscan.MultiPage1.Pages("pgEditados").ListBoxEditados
                .AddItem
                .List(ind, 0) = varArquivo
            End With

            ind = ind + 1
    
            varArquivosDir = Dir()
        Loop
        '
        '   PAG
        '
        varArquivosDir = Dir(varPasta & "ED_PAG*.xml")

        Do While varArquivosDir <> ""
            varArquivo = CStr(varArquivosDir)
            With UserForm_Zscan.MultiPage1.Pages("pgEditados").ListBoxEditados
                .AddItem
                .List(ind, 0) = varArquivo
            End With

            ind = ind + 1
    
            varArquivosDir = Dir()
        Loop
    Else
    
        '   Demais
        '
        varArquivosDir = Dir(varPasta & "ED_" & Pesquisar_chave(mTela, "ArquivoBase") & "*.XLSX")

        ind = 0
        Do While varArquivosDir <> ""
            varArquivo = CStr(varArquivosDir)

            With UserForm_Zscan.MultiPage1.Pages("pgEditados").ListBoxEditados
                .AddItem
                .List(ind, 0) = varArquivo
            End With

            ind = ind + 1
    
            varArquivosDir = Dir()
        Loop
    
    End If

End Sub

Public Sub fechar()
    ActiveWorkbook.Close (False)
End Sub

Sub CriarPastas()
    On Error Resume Next

    Dim varPasta As String
    Dim varInd As Integer

    varPasta = Pesquisar_chave("Geral", "Pasta Padrão")
    
    CriarFolder (varPasta)
    
    varInd = 6
    While Sheets("Acessos").Cells(varInd, 2) <> ""
        CriarFolder (varPasta & Sheets("Acessos").Cells(varInd, 7))
        CriarFolder (varPasta & Sheets("Acessos").Cells(varInd, 7) & "temp\")
        varInd = varInd + 1
    Wend
    
End Sub


