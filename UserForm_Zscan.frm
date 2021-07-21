VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_Zscan 
   Caption         =   "Importação "
   ClientHeight    =   7600
   ClientLeft      =   110
   ClientTop       =   2690
   ClientWidth     =   17190
   OleObjectBlob   =   "UserForm_Zscan.frx":0000
End
Attribute VB_Name = "UserForm_Zscan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "Formulários"
Option Explicit

Private Sub cbConfirmaZscan_Click()
    On Error GoTo cbConfirmaZscan_Click_Err

    Dim mReferencia As String
    
    tbTotLaudos_O = ""
    tbTotLaudos_E = ""
    tbReferencia = ""
    

    '---------------------------------------------------------------------------------------
    '
    mOperadora = UserForm_Zscan.ComboOperadora.Value

    ExecutaBotões mTela, mOperadora
    '    Importar_Arquivo mTela
    

    tbReferencia = ""
    tbTotLaudos_O = ""
    tbTotLaudos_E = ""
    cbConfirmaZscan.Enabled = True

    Exit Sub

cbConfirmaZscan_Click_Err:

    MsgBox Error$

    cbConfirmaZscan.Enabled = True
End Sub

Private Sub CommandButton1_Click()
    Dim X As Integer
    For X = 0 To ListBoxLog.ListCount - 1
        'Verifica se o item do listbox esta selecionado
        If ListBoxLog.Selected(X) Then
            'Se estiver selecionado escreve o resultado no TextBox
            Debug.Print "CB1: " & CStr(ListBoxLog.List(X, 1))
        End If
    Next

End Sub

Private Sub ComboOperadora_Change()


    Call Montar_listbox_baixados(ComboOperadora.Value)
    
    Call Montar_listbox_editados(ComboOperadora.Value)

End Sub

Private Sub ListBoxLog_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Debug.Print "db: " & CStr(ListBoxLog.List(ListBoxLog.ListIndex, 1))

End Sub

Private Sub MultiPage1_Change()

    Call Montar_listbox_baixados(ComboOperadora.Value)

End Sub

Private Sub UserForm_Initialize()

    Dim lin As Integer
 
    With Me
        .Height = Application.Height - .Top - 8
        .Width = Application.Width - .Left - 8
    End With

    With Me.MultiPage1
        .Height = Me.Height - 10
        .Width = Me.Width - 10
    End With
        
    If mTela = "Zscan" Then
    
        FrOperadora.Visible = True
        FrOperadora.Caption = "Arquivo"
        ComboOperadora.AddItem mTela
        ComboOperadora.Value = mTela

        Me.lbOriginal.Caption = "Total de Laudos original"
        Me.lbEditado.Caption = "Total de Laudos editado"
    ElseIf mTela = "Guias" Then
    
        FrOperadora.Visible = True
        FrOperadora.Caption = "Arquivo"
        ComboOperadora.AddItem mTela
        ComboOperadora.Value = mTela

        Me.lbOriginal.Caption = "Total de Guias original"
        Me.lbEditado.Caption = "Total de Guias editado"
    
    ElseIf mTela = "CASSE" Then
    
        FrOperadora.Visible = False

        Me.lbOriginal.Caption = "Total de Recibos original"
        Me.lbEditado.Caption = "Total de Recibos editado"
    
    ElseIf mTela = "Operadora" Then
    
        FrOperadora.Caption = "Operadora"
        FrOperadora.Visible = True

        lin = 6
        Do Until Sheets("Acessos").Cells(lin, 1) = ""
            If Sheets("Acessos").Cells(lin, 1) = "Operadora" Then
                ComboOperadora.AddItem Sheets("Acessos").Cells(lin, 2)
            End If
            lin = lin + 1
        Loop

        Me.lbOriginal.Caption = "Total de Recibos original"
        Me.lbEditado.Caption = "Total de Recibos editado"
    
    End If
    Montar_listbox_do_log

End Sub

