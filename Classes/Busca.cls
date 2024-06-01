VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Busca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'** IMPORTANTE PARA UTILIZAÇÃO DA CLASSE **'
'**** Toda vez que uma função recebe um vetor como parametro
'**** é necessario criar, redimencionar e alimentar esses vetores no local onde a função será chamada
'**** para depois sim passá-los como parametro


'** FUNÇÃO QUE APAGA A BUSCA DA TELA
'** Recebe como parametros
'**** NOME DA PRIMEIRA CELULA que vai receber a busca
'**** NOME DA TELA
Public Function limpaCelulasRecebeBusca(nomePrimeiraCelula As String, nomeTela As String)
    '** Seleciona a primeira celula da tabela q recebeu a busca
    Sheets(nomeTela).Range(nomePrimeiraCelula).Select
    '** Seleciona todas as linhas
    Sheets(nomeTela).Range(Selection, Selection.End(xlDown)).Select
    '** Seleciona todas as colunas
    Sheets(nomeTela).Range(Selection, Selection.End(xlToRight)).Select
    '** Apaga tds os dados selecinados
    Selection.ClearContents
    '** Reposiciona o cursor
    Range(nomePrimeiraCelula).Select

End Function



'** FUNÇÃO QUE FAZ A BUSCA DOS DADOS
'** Recebe como parametro:
'**** NUMERO DE CAMPOS de busca da tela
'**** NOME DA TELA
'**** NOME DA TABELA
'**** VETOR de lista de celulas da tela
'**** VETOR de lista de colunas da tabela
'**** VETOR que recebe o tipo de busca -- 0 para busca exata e 1 para busca de todos os resultados que contem a informação
'**** LINHA da tela que recebe a busca
'**** COLUNA da tela que recebe a busca
'**** NUMERO DE CAMPOS da tabela
Public Function buscaDados(numeroCampoBusca As Integer, nomeTela As String, nomeTabela As String, listaDeCelulas() As Variant, listaColunasTabela() As Variant, tipoBusca() As Integer, linhaTela As Integer, colunaTela As Integer, numeroCamposTabela As Integer)
    
    '** Cria e estancia um objeto tipo funcionalidades
    Dim funcionalidades As New funcionalidades
    
    '** Vetor que guarda a verificação se a celula esta vazia ou não
    Dim verificaCampoVazio() As Boolean
    ReDim verificaCampoVazio(numeroCampoBusca - 1)
    '** Vetor que guarda o valor dos campos da tela de busca que será comparado
    Dim campoTela() As String
    ReDim campoTela(numeroCampoBusca - 1)
    '** Vetor que guarda o endereço do campo da tabela que será comparado
    Dim campoTabela() As String
    ReDim campoTabela(numeroCampoBusca - 1)
    '** Vetor que guarda o tipo de comparação que será realizada. 0 para busca exata, 1 para busca que entrega todos os resultados que contem o valor que será buscado
    Dim verificaTipoBusca() As Integer
    ReDim verificaTipoBusca(numeroCampoBusca - 1)
    
    '** Variavel que guarda o numero referente a quantidade de campos na tela que comtém informação
    Dim contadorCampos As Integer
    contadorCampos = 0
    '** Variavel que guarda todo o codigo de busca que  será criado
    Dim executaCodigo As String
    '** Variavel que verifica se terá uma busca ou se os campos de busca da tela estão vazios
    Dim constroeCondicoes As Boolean
    constroeCondicoes = False
    
    
    '** Verifica quais campos tem dados
    '** Subtrai 1 da variavel numeroCamposBusca porque é informado a quantidade de campos que recebem informações de busca na tela mas os indices dos vetores começam em zero
    For contador = 0 To numeroCampoBusca - 1
        '** Verifica se o campo não está vazio
        If IsEmpty(Sheets(nomeTela).Range(listaDeCelulas(contador)).Value) = False Then
            
            '** Recebe true indicando que o campo não está vazio
            verificaCampoVazio(contadorCampos) = True
            '** Recebe o valor do campo da tela que será comparado
            campoTela(contadorCampos) = Sheets(nomeTela).Range(listaDeCelulas(contador)).Value
            '** Recebe o endereço do campo da tabela que será comparado
            campoTabela(contadorCampos) = "Sheets(""" & nomeTabela & """).Cells(contador + 1," & listaColunasTabela(contador) & ").Value"
            '** Recebe o tipo de comparação que será feito
            verificaTipoBusca(contadorCampos) = tipoBusca(contador)
            '** Incrementa o contador de campos que contem informação
            contadorCampos = contadorCampos + 1
            
            '** Recebe True indicando que terá uma busca
            constroeCondicoes = True
        End If
    Next contador
    
    '** If que verifica se terá uma busca ou se os campos estão vazios
    '** Caso os campos estiverem vazios, vai para o fim da função
    If constroeCondicoes = False Then
        GoTo Fim:
    End If
    
    '** Inicia variavel que verifica se a primeira condição ja foi montada ou não
    contaPrimeiraCondicao = False
    
    '** MONTA A NOVA SUB QUE RECEBE O CÓDIGO CONSTRUIDO **
    '** A variavel executaCodigo é uma string que recebe o códdigo
    '** vbCrLf instrução para pular linha
    
    '** Constroe a Sub
    executaCodigo = "Sub novaBusca()" & vbCrLf
    
    '** Constroe no novo codigo a variavel que verifica se o item foi encontrado
    executaCodigo = executaCodigo + "Dim verificaItemEncontrado as boolean" & vbCrLf
    '** Constroe no novo codigo a variavel linhaTela que recebe o valor real da variavel linhaTela
    executaCodigo = executaCodigo + "linhaTela = " & linhaTela & vbCrLf
    '** Constroe no novo codigo a variavel colunaTela que recebe o valor real da variavel colunaTela
    executaCodigo = executaCodigo + "colunaTela = " & colunaTela & vbCrLf
    '** Constroe no novo codigo a variavel iniciaColuna que recebe o valor real da variavel colunaTela
    executaCodigo = executaCodigo + "iniciaColuna = " & colunaTela & vbCrLf
    
    '** Inicia a variavel que verifica se o item foi encontrado
    executaCodigo = executaCodigo + "verificaItemEncontrado = false" & vbCrLf
    
    '** Encontra ultima linha da tabela
    ultimaLinha = Sheets(nomeTabela).Cells(Rows.Count, 1).End(xlUp).Row
    
    '** Constroe no novo codigo o For e passa com o valor real da ultimaLinha
    executaCodigo = executaCodigo + "   For contador = 0 To " & ultimaLinha & vbCrLf
    
     '** Constroe no novo codigo o if
    executaCodigo = executaCodigo + "       if "
    
    '** For que monta as condições da nova busca
    For contador = 0 To numeroCampoBusca - 1
    
        '** If que verifica se a condição que vai ser construida é a primeira ou não
        If verificaCampoVazio(contador) = True And contaPrimeiraCondicao = False Then
            
            '** If que verifica qual o tipo de comparação vai ser construida
            If verificaTipoBusca(contador) = 0 Then
                '** Constroe no novo codigo a comparação
                executaCodigo = executaCodigo + campoTela(contador) & " = " & campoTabela(contador)
            Else
                '** Constroe no novo codigo a comparação
                executaCodigo = executaCodigo + "InStr(1, " & campoTabela(contador) & ", """ & campoTela(contador) & """, vbTextCompare)"
            End If
            
            '** Armazena se a primeira condição foi montada
            contaPrimeiraCondicao = True
            
        ElseIf verificaCampoVazio(contador) = True Then
            
            '** If que verifica qual o tipo de comparação vai ser construida
            If verificaTipoBusca(contador) = 0 Then
                '** Constroe no novo codigo a comparação
                executaCodigo = executaCodigo + " and " & campoTela(contador) & " = " & campoTabela(contador)
            Else
                '** Constroe no novo codigo a comparação
                executaCodigo = executaCodigo + " and InStr(1, " & campoTabela(contador) & ", """ & campoTela(contador) & """, vbTextCompare)"
            End If
            
        End If
    
    Next contador
    
    '** continua montagem do codigo
    executaCodigo = executaCodigo + " then" & vbCrLf
    
    
    '** Constroe no novo codigo
    '**** O if que verifica se a primeira linha encontrada é a linha de titulo da tabela. Caso for ela é desconsiderada
    executaCodigo = executaCodigo + "           If Sheets(""" & nomeTabela & """).Cells(contador + 1, 1).Value = ""ID"" then " & vbCrLf
    executaCodigo = executaCodigo + "               goto FimDoLoop:" & vbCrLf
    executaCodigo = executaCodigo + "           else" & vbCrLf
    executaCodigo = executaCodigo + "               verificaItemEncontrado = True" & vbCrLf
    
    '** Constroe no novo codigo o For que percorre os registros da tabela
    executaCodigo = executaCodigo + "               For contadorColunaTabela = 1 To " & numeroCamposTabela & vbCrLf
    
    '** Joga os dados na tela de busca
    executaCodigo = executaCodigo + "                   Sheets(""" & nomeTela & """).Cells(linhaTela, colunaTela) = Sheets(""" & nomeTabela & """).Cells(contador + 1, contadorColunaTabela).Value" & vbCrLf
    '** Incrementa a variavel colunaTela
    executaCodigo = executaCodigo + "                   colunaTela = colunaTela + 1" & vbCrLf
    
    
    executaCodigo = executaCodigo + "               Next contadorColunaTabela" & vbCrLf
    executaCodigo = executaCodigo + "           end if" & vbCrLf
    
    '** Muda a linha da tela q vai receber a busca
    executaCodigo = executaCodigo + "           linhaTela = linhaTela + 1" & vbCrLf
    '** Volta a colunaTela ao inicio
    executaCodigo = executaCodigo + "           colunaTela = iniciaColuna" & vbCrLf
    executaCodigo = executaCodigo + "           FimDoLoop:" & vbCrLf
    '** Fim do If
    executaCodigo = executaCodigo + "       end if" & vbCrLf
    '** Fim do For
    executaCodigo = executaCodigo + "   Next contador" & vbCrLf
    
    '** Constroe no novo codigo o IF que verifa se a busca encontrou algum dado
    executaCodigo = executaCodigo + "   If verificaItemEncontrado = False Then" & vbCrLf
    executaCodigo = executaCodigo + "       Msgbox ""Registro Não Encontrado""" & vbCrLf
    executaCodigo = executaCodigo + "   end if" & vbCrLf
    
    '** Apaga a nova SUB criada
    executaCodigo = executaCodigo + "   ThisWorkbook.VBProject.VBComponents.Remove ThisWorkbook.VBProject.VBComponents(""BuscaTemporaria"")" & vbCrLf
    '** Finaliza a o codigo
    executaCodigo = executaCodigo + "End Sub"
      
    '** Criar um novo módulo
    Dim novoModulo As Object
    
    '** 1 significa vbext_ct_StdModule, que é um módulo padrão
    Set novoModulo = ThisWorkbook.VBProject.VBComponents.Add(1)
    
    '** Renomeia o Modulo
    novoModulo.Name = "BuscaTemporaria"
    
    '** Adicionar código ao novo módulo
    novoModulo.CodeModule.AddFromString executaCodigo

    '** Executar a nova sub-rotina
    Application.Run "novaBusca"
    
    '** Chama a função que apaga os campos da Tela
    funcionalidades.apagaCampos nomeTela, listaDeCelulas()
    Exit Function
        
'** Marca a chamada do Fim da função
Fim:
        MsgBox "Não existem dados nos Campos de Busca"
    
End Function

