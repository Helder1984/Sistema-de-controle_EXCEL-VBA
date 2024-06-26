VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Cadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'** IMPORTANTE PARA UTILIZA��O DA CLASSE **'
'**** Toda vez que uma fun��o recebe um vetor como parametro
'**** � necessario criar, redimencionar e alimentar esses vetores no local onde a fun��o ser� chamada
'**** para depois sim pass�-los como parametro



'** Fun��o q atribui novo ID automaticamente a uma TELA DE CADASTRO
'** Geralmente chamada no Worksheet_SelectionChange da Planilha da tela
'** Recebe como Parametros
'**** NOME DA TABELA
'**** NOME DA TELA
'**** CELULA da tela que receber� o novo c�digo
Public Function novoID(nomeTabela As String, nomeTela As String, ByVal CelulaRecebeCodigo As String)
    
    '**Encontra a �ltima linha da Tabela
    '**e atribui o numero da linha na vari�vel
    ultimaLinha = Sheets(nomeTabela).Cells(Rows.Count, 1).End(xlUp).Row
    '**Envia novo ID para a tela
    Sheets(nomeTela).Range(CelulaRecebeCodigo) = Sheets(nomeTabela).Range("A" & ultimaLinha).Value + 1

End Function


'** FUN��O QUE SALVA OS DADOS NA TABELA
'** Recebe como parametro
'**** NOME DA TELA
'**** VETOR que recebe a lista de nome dos campos
'**** NOME DA TABELA
'**** VAREAVEL DE VALIDA�AO q verifica se todos os campos passaram ou n�o nos testes
Public Function salvaDadosTabela(nomeTela As String, listaDeCelulas() As Variant, nomeTabela As String, validacao As Boolean)
    
    '** If que verifica se todos os campos passaram na valida��o
    '** Caso n�o passarem aborta a fun��o
    If validacao = False Then
        Exit Function
    End If
    
    '** Cria e estancia um objeto tipo funcionalidades
    Dim funcionalidades As New funcionalidades
    
    '**Encontra a �ltima linha da Tabela
    '**e "adiciona" + 1 Linha
    adicionaLinhaTabela = Sheets(nomeTabela).Cells(Rows.Count, 1).End(xlUp).Row + 1
    
    '** Contador das colunas da tabela
    '** N�o pode ser igual ao contador porque o indice de colunas da tabela come�a em 1
    numeroColuna = 1
    
    '** Loop q pega os dados da tela e joga para a tabela
    '** O UBound � onde acaba o vetor
    For contador = 0 To UBound(listaDeCelulas)
        '** Campo da tabela recebe o valor do campo da tela
        Sheets(nomeTabela).Cells(adicionaLinhaTabela, numeroColuna) = (Sheets(nomeTela).Range(listaDeCelulas(contador)).Value)
        '** Incrementa o contador das colunas da tabela
        numeroColuna = numeroColuna + 1
    Next contador
    
    '** Mensagem de sucesso
    MsgBox ("Cadastro Realizado com sucesso")
    
    '** Chama fun��o que limpa os campos da tela
    '** Passa como parametro
    '**** NOME DA TELA
    '**** VETOR com a lista de nome dos campos
    funcionalidades.apagaCampos nomeTela, listaDeCelulas
    
    '** Chama a fun��o q adiciona um novo ID na tela
    '** Passa como parametro
    '**** NOME DA TABELA
    '**** NOME DA TELA
    '**** VETOR de nome dos campos na posi��o 0 que sempre se refere ao ID
    novoID nomeTabela, nomeTela, listaDeCelulas(0)
    
    '** Posiciona o cursor no primeiro campo a ser preenchido
    Sheets(nomeTela).Range(listaDeCelulas(1)).Select

End Function
