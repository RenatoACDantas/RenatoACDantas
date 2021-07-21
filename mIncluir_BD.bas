Attribute VB_Name = "mIncluir_BD"
'@Folder "Importa��o"
Option Explicit

Sub Incluir_BD(Se��o As String)
    On Error GoTo Incluir_BD_Err

    Dim mCaminho As String

    Dim obj As Object

    Dim appObj As Object

    mCaminho = Pesquisar_chave("Geral", "Pasta Padr�o") & Pesquisar_chave("Geral", "BD")

    mTabela = Pesquisar_chave(Se��o, "TabelaAccess")
    mRange = Pesquisar_chave(Se��o, "Range")

    Montar_listbox "Inclus�o BD: ", mTabela

    Set appObj = CreateObject("Access.Application")

    appObj.OpenCurrentDatabase mCaminho

    appObj.Run "importarPlanilha", mTabela, mArquivoSaida, mRange, "'" + mReferencia + "'"

    appObj.Quit
    Application.Cursor = xlDefault
    Montar_listbox "T�rmimo inclus�o BD: ", mTela
    Montar_listbox "-----", "-----"

Incluir_BD_Exit:
    Exit Sub

Incluir_BD_Err:
    Application.Cursor = xlDefault
    Montar_listbox "Erro inclus�o BD: ", Err.Description

    MsgBox Error$

End Sub

Sub t()
    Application.Cursor = xlDefault
End Sub

