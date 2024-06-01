VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Edicao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'** IMPORTANTE PARA UTILIZAÇÃO DA CLASSE **'
'**** Toda vez que uma função recebe um vetor como parametro
'**** é necessario criar, redimencionar e alimentar esses vetores no local onde a função será chamada
'**** para depois sim passá-los como parametro



'** FUNÇÃO QUE ENCONTRA O REGISTRO QUE SERÁ ALTERADO
'** Recebe como parametro
'**** NUMERO DA COLUNA da tela de busca que se refere ao ID
'**** NOME DA TABELA
'**** NOME DA TELA DE BUSCA
'**** NOME DA TELA DE EDIÇÃO
'**** VETOR que recebe o nome dos campos da tela de edição
Public Function encontraRegistro(colunaTela As Integer, nomeTabela As String, nomeTelaBusca As String, nomeTelaEdicao As String, listaDeCelulasTelaEdicao() As Variant)
          
    '** Declara uma variável do tipo Range para armazenar a seleção atual
    Dim celulaSelecionada As Range
      
    '** Verifica se a seleção atual é uma célula ou um intervalo de células
    If TypeName(Selection) = "Range" Then
        '** Atribui a célula ou intervalo selecionado à variável
        Set celulaSelecionada = Selection
        
        '** Verifica se a seleção está apenas em uma linha
        If celulaSelecionada.Rows.Count = 1 Then
            '** Atribui a variavel a linhaTela o numero da linha selecionada selecionada
            linhaTela = celulaSelecionada.Row
            
            '** Pula direto para a continuação da edição
            GoTo Editar:
        Else
            '** Exibe uma mensagem caso mais de uma linha seja selecionada
            MsgBox "Selecione apenas um item", vbInformation
            
            '** Pula direto para o fim da função
            GoTo Fim:
        End If
    Else
        ' Exibe uma mensagem se a seleção não for uma célula ou intervalo de células
        MsgBox "A seleção atual não é um produto valido", vbExclamation
    End If
    
Editar:

    '** Encontra ultima linha da tabela
    ultimaLinha = Sheets(nomeTabela).Cells(Rows.Count, 1).End(xlUp).Row
    
    '** Loop que encontra o registro selecionado
    '** contador = 1 porque as colunas começam em 1
    For contador = 1 To ultimaLinha
        
        '** If que compara o item selecionado com os registros da tabela
        If Sheets(nomeTelaBusca).Cells(linhaTela, colunaTela) = Sheets(nomeTabela).Cells(contador, 1) Then
            
            '** Loop que alimenta a tela de edição com o registro encontrado
            '** contadorMostraTela começa em 0 pq os vetores começam em 0
            For contadorMostraTela = 0 To UBound(listaDeCelulasTelaEdicao)
            
                '** Tela de edição recebe os dados da tabela
                Sheets(nomeTelaEdicao).Range(listaDeCelulasTelaEdicao(contadorMostraTela)) = Sheets(nomeTabela).Cells(contador, contadorMostraTela + 1)
            
            Next contadorMostraTela
            
            '** Vai para a chamada da tela de edição
            GoTo TelaEdicao:
        End If
        
    Next contador
    
    '** Se não for encontrado nenhum registo
    '** Exibe a mensagem e vai para o fim da função
    MsgBox ("Selecione algum item")
    GoTo Fim:
    
TelaEdicao:
    '** Chama a tela de edição
    ThisWorkbook.Sheets(nomeTelaEdicao).Activate
    
Fim:
    
End Function

'** FUNÇÃO QUE SALVA AS ALTERAÇÕES
'** Recebe como parametro
'**** NOME DA TELA de edição
'**** VETOR com os nomes dos campos da tela de edição
'**** NOME DA TABELA
'**** VARIAVEL que controla se os campos passaram na validação
Public Function Editar(nomeTela As String, listaDeCelulas() As Variant, nomeTabela As String, validacao As Boolean)
    
    '** If que verifica se os campos passaram na validação
    '** Caso não passaram, encerra o codigo
    If validacao = False Then
        End
    End If
    
    '** Confirma se deseja ou não salvar as alterações
    '** Caso não deseje, encerra o codigo
    Dim resposta As VbMsgBoxResult
    resposta = MsgBox("Deseja salvar as alterações?", vbYesNo + vbQuestion, "Confirmação")
    '** If de controle das ações
    If resposta = vbYes Then
    Else
        End
    End If
    
    '** Encontra ultima linha da tabela
    ultimaLinha = Sheets(nomeTabela).Cells(Rows.Count, 1).End(xlUp).Row
    
    '** Loop que encontra na tabela o item que será alterado
    For contador = 1 To ultimaLinha
    
        '** If que verifica se o id do item da tela é o mesmo do item da tabela
        If Sheets(nomeTela).Range(listaDeCelulas(0)) = Sheets(nomeTabela).Cells(contador, 1) Then
            '** Soma 1 ao contador para percorrer até o fim do vetor
            fimContador = UBound(listaDeCelulas) + 1
            
            '** Loop que salva as alterações na tabela
            '** O contador começa no 2 para garantir que o ID não seja modificado
            For contadorSalvaDados = 2 To fimContador
                Sheets(nomeTabela).Cells(contador, contadorSalvaDados) = Sheets(nomeTela).Range(listaDeCelulas(contadorSalvaDados - 1))
            Next contadorSalvaDados
            
        End If
        
    Next contador
    
    '** Mensagem de confirmação
    MsgBox "Dados alterados com sucesso"
    
End Function

