VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Edicao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'** IMPORTANTE PARA UTILIZA��O DA CLASSE **'
'**** Toda vez que uma fun��o recebe um vetor como parametro
'**** � necessario criar, redimencionar e alimentar esses vetores no local onde a fun��o ser� chamada
'**** para depois sim pass�-los como parametro



'** FUN��O QUE ENCONTRA O REGISTRO QUE SER� ALTERADO
'** Recebe como parametro
'**** NUMERO DA COLUNA da tela de busca que se refere ao ID
'**** NOME DA TABELA
'**** NOME DA TELA DE BUSCA
'**** NOME DA TELA DE EDI��O
'**** VETOR que recebe o nome dos campos da tela de edi��o
Public Function encontraRegistro(colunaTela As Integer, nomeTabela As String, nomeTelaBusca As String, nomeTelaEdicao As String, listaDeCelulasTelaEdicao() As Variant)
          
    '** Declara uma vari�vel do tipo Range para armazenar a sele��o atual
    Dim celulaSelecionada As Range
      
    '** Verifica se a sele��o atual � uma c�lula ou um intervalo de c�lulas
    If TypeName(Selection) = "Range" Then
        '** Atribui a c�lula ou intervalo selecionado � vari�vel
        Set celulaSelecionada = Selection
        
        '** Verifica se a sele��o est� apenas em uma linha
        If celulaSelecionada.Rows.Count = 1 Then
            '** Atribui a variavel a linhaTela o numero da linha selecionada selecionada
            linhaTela = celulaSelecionada.Row
            
            '** Pula direto para a continua��o da edi��o
            GoTo Editar:
        Else
            '** Exibe uma mensagem caso mais de uma linha seja selecionada
            MsgBox "Selecione apenas um item", vbInformation
            
            '** Pula direto para o fim da fun��o
            GoTo Fim:
        End If
    Else
        ' Exibe uma mensagem se a sele��o n�o for uma c�lula ou intervalo de c�lulas
        MsgBox "A sele��o atual n�o � um produto valido", vbExclamation
    End If
    
Editar:

    '** Encontra ultima linha da tabela
    ultimaLinha = Sheets(nomeTabela).Cells(Rows.Count, 1).End(xlUp).Row
    
    '** Loop que encontra o registro selecionado
    '** contador = 1 porque as colunas come�am em 1
    For contador = 1 To ultimaLinha
        
        '** If que compara o item selecionado com os registros da tabela
        If Sheets(nomeTelaBusca).Cells(linhaTela, colunaTela) = Sheets(nomeTabela).Cells(contador, 1) Then
            
            '** Loop que alimenta a tela de edi��o com o registro encontrado
            '** contadorMostraTela come�a em 0 pq os vetores come�am em 0
            For contadorMostraTela = 0 To UBound(listaDeCelulasTelaEdicao)
            
                '** Tela de edi��o recebe os dados da tabela
                Sheets(nomeTelaEdicao).Range(listaDeCelulasTelaEdicao(contadorMostraTela)) = Sheets(nomeTabela).Cells(contador, contadorMostraTela + 1)
            
            Next contadorMostraTela
            
            '** Vai para a chamada da tela de edi��o
            GoTo TelaEdicao:
        End If
        
    Next contador
    
    '** Se n�o for encontrado nenhum registo
    '** Exibe a mensagem e vai para o fim da fun��o
    MsgBox ("Selecione algum item")
    GoTo Fim:
    
TelaEdicao:
    '** Chama a tela de edi��o
    ThisWorkbook.Sheets(nomeTelaEdicao).Activate
    
Fim:
    
End Function

'** FUN��O QUE SALVA AS ALTERA��ES
'** Recebe como parametro
'**** NOME DA TELA de edi��o
'**** VETOR com os nomes dos campos da tela de edi��o
'**** NOME DA TABELA
'**** VARIAVEL que controla se os campos passaram na valida��o
Public Function Editar(nomeTela As String, listaDeCelulas() As Variant, nomeTabela As String, validacao As Boolean)
    
    '** If que verifica se os campos passaram na valida��o
    '** Caso n�o passaram, encerra o codigo
    If validacao = False Then
        End
    End If
    
    '** Confirma se deseja ou n�o salvar as altera��es
    '** Caso n�o deseje, encerra o codigo
    Dim resposta As VbMsgBoxResult
    resposta = MsgBox("Deseja salvar as altera��es?", vbYesNo + vbQuestion, "Confirma��o")
    '** If de controle das a��es
    If resposta = vbYes Then
    Else
        End
    End If
    
    '** Encontra ultima linha da tabela
    ultimaLinha = Sheets(nomeTabela).Cells(Rows.Count, 1).End(xlUp).Row
    
    '** Loop que encontra na tabela o item que ser� alterado
    For contador = 1 To ultimaLinha
    
        '** If que verifica se o id do item da tela � o mesmo do item da tabela
        If Sheets(nomeTela).Range(listaDeCelulas(0)) = Sheets(nomeTabela).Cells(contador, 1) Then
            '** Soma 1 ao contador para percorrer at� o fim do vetor
            fimContador = UBound(listaDeCelulas) + 1
            
            '** Loop que salva as altera��es na tabela
            '** O contador come�a no 2 para garantir que o ID n�o seja modificado
            For contadorSalvaDados = 2 To fimContador
                Sheets(nomeTabela).Cells(contador, contadorSalvaDados) = Sheets(nomeTela).Range(listaDeCelulas(contadorSalvaDados - 1))
            Next contadorSalvaDados
            
        End If
        
    Next contador
    
    '** Mensagem de confirma��o
    MsgBox "Dados alterados com sucesso"
    
End Function

