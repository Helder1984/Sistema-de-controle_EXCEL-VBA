Attribute VB_Name = "AjustarTela"
Sub AjustaTela()
Attribute AjustaTela.VB_ProcData.VB_Invoke_Func = " \n14"

    'Tratamento de erro
    On Error GoTo Sair
    
    'Seleciona o intervalo que definimos
    ActiveSheet.Range("Tela").Select
    
    'Ajusta o zoom da tela para melhor configura��o possivel
    ActiveWindow.Zoom = True
    
    'Volta a sele��o para a celula A1
    'Apenas est�tica
    ActiveSheet.Range("A1").Select
    
Sair:
    Exit Sub
End Sub
