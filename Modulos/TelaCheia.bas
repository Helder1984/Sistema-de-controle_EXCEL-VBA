Attribute VB_Name = "TelaCheia"
Sub Sair()
Attribute Sair.VB_ProcData.VB_Invoke_Func = " \n14"
    'Fecha o Excel
    Application.Quit
End Sub


Sub AtivaTelaCheia()
        
    'Oculta as guias de menu
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"
    
    'Ocultar barra de f�rmulas
    Application.DisplayFormulaBar = False
    
    'Ocultar barra de status, a barrinha cinza l� embaixo
    Application.DisplayStatusBar = False
    
    With ActiveWindow
        'Ocultar barra horizontal
        .DisplayHorizontalScrollBar = False
        
        'Ocultar barra vertical
        .DisplayVerticalScrollBar = False
        
        'Ocultar abas das planilhas
        .DisplayWorkbookTabs = False
        
        'Oculta os t�tulos
        .DisplayHeadings = False

        'Oculta as linhas de grade da planilha
        .DisplayGridlines = False
    End With
    
    AjustaTela
    
End Sub

Sub DesligaTelaCheia()
    
    'Exibe as guias de menu
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
    
    'Exibe as barra de f�rmulas
    Application.DisplayFormulaBar = True
    
    'Exibe barra de status, a barrinha cinza l� embaixo
    Application.DisplayStatusBar = True
    
    With ActiveWindow
        'Exibe barra horizontal
        .DisplayHorizontalScrollBar = True
        
        'Exibe barra vertical
        .DisplayVerticalScrollBar = True
        
        'Exibe guias das planilhas
        .DisplayWorkbookTabs = True
        
        'Exibe os t�tulos de linha e coluna
        .DisplayHeadings = True
        
        'Exibe as linhas de grade da planilha
        .DisplayGridlines = True
    End With
    
End Sub
