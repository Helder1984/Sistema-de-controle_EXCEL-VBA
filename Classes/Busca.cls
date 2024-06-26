VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Busca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'** IMPORTANTE PARA UTILIZA��O DA CLASSE **'
'**** Toda vez que uma fun��o recebe um vetor como parametro
'**** � necessario criar, redimencionar e alimentar esses vetores no local onde a fun��o ser� chamada
'**** para depois sim pass�-los como parametro


'** FUN��O QUE APAGA A BUSCA DA TELA
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



'** FUN��O QUE FAZ A BUSCA DOS DADOS
'** Recebe como parametro:
'**** NUMERO DE CAMPOS de busca da tela
'**** NOME DA TELA
'**** NOME DA TABELA
'**** VETOR de lista de celulas da tela
'**** VETOR de lista de colunas da tabela
'**** VETOR que recebe o tipo de busca -- 0 para busca exata e 1 para busca de todos os resultados que contem a informa��o
'**** LINHA da tela que recebe a busca
'**** COLUNA da tela que recebe a busca
'**** NUMERO DE CAMPOS da tabela
Public Function buscaDados(numeroCampoBusca As Integer, nomeTela As String, nomeTabela As String, listaDeCelulas() As Variant, listaColunasTabela() As Variant, tipoBusca() As Integer, linhaTela As Integer, colunaTela As Integer, numeroCamposTabela As Integer)
    
    '** Cria e estancia um objeto tipo funcionalidades
    Dim funcionalidades As New funcionalidades
    
    '** Vetor que guarda a verifica��o se a celula esta vazia ou n�o
    Dim verificaCampoVazio() As Boolean
    ReDim verificaCampoVazio(numeroCampoBusca - 1)
    '** Vetor que guarda o valor dos campos da tela de busca que ser� comparado
    Dim campoTela() As String
    ReDim campoTela(numeroCampoBusca - 1)
    '** Vetor que guarda o endere�o do campo da tabela que ser� comparado
    Dim campoTabela() As String
    ReDim campoTabela(numeroCampoBusca - 1)
    '** Vetor que guarda o tipo de compara��o que ser� realizada. 0 para busca exata, 1 para busca que entrega todos os resultados que contem o valor que ser� buscado
    Dim verificaTipoBusca() As Integer
    ReDim verificaTipoBusca(numeroCampoBusca - 1)
    
    '** Variavel que guarda o numero referente a quantidade de campos na tela que comt�m informa��o
    Dim contadorCampos As Integer
    contadorCampos = 0
    '** Variavel que guarda todo o codigo de busca que  ser� criado
    Dim executaCodigo As String
    '** Variavel que verifica se ter� uma busca ou se os campos de busca da tela est�o vazios
    Dim constroeCondicoes As Boolean
    constroeCondicoes = False
    
    
    '** Verifica quais campos tem dados
    '** Subtrai 1 da variavel numeroCamposBusca porque � informado a quantidade de campos que recebem informa��es de busca na tela mas os indices dos vetores come�am em zero
    For contador = 0 To numeroCampoBusca - 1
        '** Verifica se o campo n�o est� vazio
        If IsEmpty(Sheets(nomeTela).Range(listaDeCelulas(contador)).Value) = False Then
            
            '** Recebe true indicando que o campo n�o est� vazio
            verificaCampoVazio(contadorCampos) = True
            '** Recebe o valor do campo da tela que ser� comparado
            campoTela(contadorCampos) = Sheets(nomeTela).Range(listaDeCelulas(contador)).Value
            '** Recebe o endere�o do campo da tabela que ser� comparado
            campoTabela(contadorCampos) = "Sheets(""" & nomeTabela & """).Cells(contador + 1," & listaColunasTabela(contador) & ").Value"
            '** Recebe o tipo de compara��o que ser� feito
            verificaTipoBusca(contadorCampos) = tipoBusca(contador)
            '** Incrementa o contador de campos que contem informa��o
            contadorCampos = contadorCampos + 1
            
            '** Recebe True indicando que ter� uma busca
            constroeCondicoes = True
        End If
    Next contador
    
    '** If que verifica se ter� uma busca ou se os campos est�o vazios
    '** Caso os campos estiverem vazios, vai para o fim da fun��o
    If constroeCondicoes = False Then
        GoTo Fim:
    End If
    
    '** Inicia variavel que verifica se a primeira condi��o ja foi montada ou n�o
    contaPrimeiraCondicao = False
    
    '** MONTA A NOVA SUB QUE RECEBE O C�DIGO CONSTRUIDO **
    '** A variavel executaCodigo � uma string que recebe o c�ddigo
    '** vbCrLf instru��o para pular linha
    
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
    
    '** For que monta as condi��es da nova busca
    For contador = 0 To numeroCampoBusca - 1
    
        '** If que verifica se a condi��o que vai ser construida � a primeira ou n�o
        If verificaCampoVazio(contador) = True And contaPrimeiraCondicao = False Then
            
            '** If que verifica qual o tipo de compara��o vai ser construida
            If verificaTipoBusca(contador) = 0 Then
                '** Constroe no novo codigo a compara��o
                executaCodigo = executaCodigo + campoTela(contador) & " = " & campoTabela(contador)
            Else
                '** Constroe no novo codigo a compara��o
                executaCodigo = executaCodigo + "InStr(1, " & campoTabela(contador) & ", """ & campoTela(contador) & """, vbTextCompare)"
            End If
            
            '** Armazena se a primeira condi��o foi montada
            contaPrimeiraCondicao = True
            
        ElseIf verificaCampoVazio(contador) = True Then
            
            '** If que verifica qual o tipo de compara��o vai ser construida
            If verificaTipoBusca(contador) = 0 Then
                '** Constroe no novo codigo a compara��o
                executaCodigo = executaCodigo + " and " & campoTela(contador) & " = " & campoTabela(contador)
            Else
                '** Constroe no novo codigo a compara��o
                executaCodigo = executaCodigo + " and InStr(1, " & campoTabela(contador) & ", """ & campoTela(contador) & """, vbTextCompare)"
            End If
            
        End If
    
    Next contador
    
    '** continua montagem do codigo
    executaCodigo = executaCodigo + " then" & vbCrLf
    
    
    '** Constroe no novo codigo
    '**** O if que verifica se a primeira linha encontrada � a linha de titulo da tabela. Caso for ela � desconsiderada
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
    executaCodigo = executaCodigo + "       Msgbox ""Registro N�o Encontrado""" & vbCrLf
    executaCodigo = executaCodigo + "   end if" & vbCrLf
    
    '** Apaga a nova SUB criada
    executaCodigo = executaCodigo + "   ThisWorkbook.VBProject.VBComponents.Remove ThisWorkbook.VBProject.VBComponents(""BuscaTemporaria"")" & vbCrLf
    '** Finaliza a o codigo
    executaCodigo = executaCodigo + "End Sub"
      
    '** Criar um novo m�dulo
    Dim novoModulo As Object
    
    '** 1 significa vbext_ct_StdModule, que � um m�dulo padr�o
    Set novoModulo = ThisWorkbook.VBProject.VBComponents.Add(1)
    
    '** Renomeia o Modulo
    novoModulo.Name = "BuscaTemporaria"
    
    '** Adicionar c�digo ao novo m�dulo
    novoModulo.CodeModule.AddFromString executaCodigo

    '** Executar a nova sub-rotina
    Application.Run "novaBusca"
    
    '** Chama a fun��o que apaga os campos da Tela
    funcionalidades.apagaCampos nomeTela, listaDeCelulas()
    Exit Function
        
'** Marca a chamada do Fim da fun��o
Fim:
        MsgBox "N�o existem dados nos Campos de Busca"
    
End Function

