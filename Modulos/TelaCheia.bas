Attribute VB_Name = "TelaCheia"
Sub Sair()
Attribute Sair.VB_ProcData.VB_Invoke_Func = " \n14"
    'Fecha o Excel
    Application.Quit
End Sub


Sub AtivaTelaCheia()
        
    'Oculta as guias de menu
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"
    
    'Ocultar barra de fórmulas
    Application.DisplayFormulaBar = False
    
    'Ocultar barra de status, a barrinha cinza lá embaixo
    Application.DisplayStatusBar = False
    
    With ActiveWindow
        'Ocultar barra horizontal
        .DisplayHorizontalScrollBar = False
        
        'Ocultar barra vertical
        .DisplayVerticalScrollBar = False
        
        'Ocultar abas das planilhas
        .DisplayWorkbookTabs = False
        
        'Oculta os títulos
        .DisplayHeadings = False

        'Oculta as linhas de grade da planilha
        .DisplayGridlines = False
    End With
    
    AjustaTela
    
End Sub

Sub DesligaTelaCheia()
    
    'Exibe as guias de menu
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
    
    'Exibe as barra de fórmulas
    Application.DisplayFormulaBar = True
    
    'Exibe barra de status, a barrinha cinza lá embaixo
    Application.DisplayStatusBar = True
    
    With ActiveWindow
        'Exibe barra horizontal
        .DisplayHorizontalScrollBar = True
        
        'Exibe barra vertical
        .DisplayVerticalScrollBar = True
        
        'Exibe guias das planilhas
        .DisplayWorkbookTabs = True
        
        'Exibe os títulos de linha e coluna
        .DisplayHeadings = True
        
        'Exibe as linhas de grade da planilha
        .DisplayGridlines = True
    End With
    
End Sub
