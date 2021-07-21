'@Folder "Planilhas"
Option Explicit

Private Sub Workbook_BeforeClose(Cancel As Boolean)

    Call Sair
    Call CloseBook

End Sub

Private Sub Workbook_Open()

    Sheets("Menu").Select

    Inicializar_Variaveis
    Ler_Configuração

    mPastaPadrao = Pesquisar_chave("Geral", "Pasta Padrão")
    Debug.Print mPastaPadrao
    mInd = 0
    mIndLog = 1
    mIndPlanilha = 1

    Call Desabilitar
    Call usuario
    Application.WindowState = xlMaximized
End Sub
