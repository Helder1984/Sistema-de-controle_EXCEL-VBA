VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Funcionalidades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'** IMPORTANTE PARA UTILIZA��O DA CLASSE **'
'**** Toda vez que uma fun��o recebe um vetor como parametro
'**** � necessario criar, redimencionar e alimentar esses vetores no local onde a fun��o ser� chamada
'**** para depois sim pass�-los como parametro


'** FUN��O QUE APAGA OS CAMPOS DA TELA
'** Recebe como parametros
'**** NOME DA TELA
'**** VETOR com a lista de NOME DOS CAMPOS
Public Function apagaCampos(nomeTela As String, listaDeCelulas())
        
        '** Loop q percorre o vetor
        '** O UBound � onde acaba o vetor
        For contador = 0 To UBound(listaDeCelulas)
            '** Apaga os dados dos campos
            Sheets(nomeTela).Range(listaDeCelulas(contador)).ClearContents
        
        Next contador

End Function

'** FUN��O QUE CRIA LISTBOX
'** Recebe como parametros
'**** NOME DA PLANILHA TABELA que contem os dados que ser�o verificados
'**** NOME DA TABELA que contem os dados que ser�o verificados
'**** NOME DA TELA onde a LISTBOX sera criada
'**** NOME DA CELULA que ser� a LISTBOX
'**** NUMERO DA COLUNA da tabela de dados que ser� usada para popular a LISTBOX
Public Function criaListBox(nomePlanilhaTabela As String, nomeTabela As String, nomeTelaCriarListbox As String, celulaTelaCriarListbox As String, colunaTabelaPopulaListbox As Integer)
    
    '** Cria uma variavel do tipo tabela
    Dim tabela As ListObject
    
    '** Cria variaveis do tipo Range
    Dim rangeTabela As Range
    Dim rangeCelula As Range
    Dim cel As Range
    
    '** Cria uma variavel do tipo String
    Dim listaItens As String
    
    '** Desprotege a Tela
    ThisWorkbook.Sheets(nomeTelaCriarListbox).Unprotect
    
    '** Define a planilha e a tabela
    Set tabela = ThisWorkbook.Worksheets(nomePlanilhaTabela).ListObjects(nomeTabela)
    
    '** Define a c�lula onde a LISTBOX ser� criada
    Set rangeCelula = ThisWorkbook.Worksheets(nomeTelaCriarListbox).Range(celulaTelaCriarListbox)
    
    '** Obtem o intervalo de dados da tabela
    Set rangeTabela = tabela.DataBodyRange
    
    '** Constroe a string com os itens da tabela separados por v�rgulas
    listaItens = ""
    For Each cel In rangeTabela.Columns(colunaTabelaPopulaListbox).Cells
        listaItens = listaItens & cel.Value & ","
    Next cel
    
    '** Remover a �ltima v�rgula
    If Len(listaItens) > 0 Then
        listaItens = Left(listaItens, Len(listaItens) - 1)
    End If
    
    '** Adicionando valida��o de dados para criar a lista suspensa
    With rangeCelula.Validation
        .Delete ' Remove qualquer valida��o existente
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:=listaItens
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With
    
    '** Protege a Tela
    ThisWorkbook.Sheets(nomeTelaCriarListbox).Protect
    
    
End Function
