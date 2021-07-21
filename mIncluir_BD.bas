Attribute VB_Name = "mIncluir_BD"
'@Folder "Importação"
Option Explicit

Sub Incluir_BD(Seção As String)
    On Error GoTo Incluir_BD_Err

    Dim mCaminho As String

    Dim obj As Object

    Dim appObj As Object

    mCaminho = Pesquisar_chave("Geral", "Pasta Padrão") & Pesquisar_chave("Geral", "BD")

    mTabela = Pesquisar_chave(Seção, "TabelaAccess")
    mRange = Pesquisar_chave(Seção, "Range")

    Montar_listbox "Inclusão BD: ", mTabela

    Set appObj = CreateObject("Access.Application")

    appObj.OpenCurrentDatabase mCaminho

    appObj.Run "importarPlanilha", mTabela, mArquivoSaida, mRange, "'" + mReferencia + "'"

    appObj.Quit
    Application.Cursor = xlDefault
    Montar_listbox "Térmimo inclusão BD: ", mTela
    Montar_listbox "-----", "-----"

Incluir_BD_Exit:
    Exit Sub

Incluir_BD_Err:
    Application.Cursor = xlDefault
    Montar_listbox "Erro inclusão BD: ", Err.Description

    MsgBox Error$

End Sub

Sub t()
    Application.Cursor = xlDefault
End Sub

