VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Configura��es"
   ClientHeight    =   3350
   ClientLeft      =   6120
   ClientTop       =   8460.001
   ClientWidth     =   7590
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "Formul�rios"
Option Explicit

Private Sub CommandButton1_Click()
    Sheets("Configura��o").Cells(8, 2) = tbPasta
    Sheets("Configura��o").Cells(13, 2) = tbModelo
    Sheets("Configura��o").Cells(14, 2) = tbLogo
    Sheets("Configura��o").Cells(15, 2) = tbBase
    mLogo = tbPasta + tbLogo
    ActiveWorkbook.Save
End Sub

Private Sub UserForm_Initialize()
    tbPasta = Sheets("Configura��o").Cells(8, 2)
    tbLogo = Sheets("Configura��o").Cells(14, 2)
    tbModelo = Sheets("Configura��o").Cells(13, 2)
    tbBase = Sheets("Configura��o").Cells(15, 2)
    mLogo = tbPasta + tbLogo
End Sub

