VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Validacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'** IMPORTANTE PARA UTILIZAÇÃO DA CLASSE **'
'**** Toda vez que uma função recebe um vetor como parametro
'**** é necessario criar, redimencionar e alimentar esses vetores no local onde a função será chamada
'**** para depois sim passá-los como parametro


'** Variavel global da Classe que controla se passou pela validação
Public confirmaValidacao As Boolean

'** FUNÇÃO CRIADA PARA INICIAR A VARIAVEL DE VALIDAÇÕES
Public Function iniciaValidacao()
    '** Variavel que controla se passou pela validação
    '** Inicia a variavel
    confirmaValidacao = True
End Function


'** FUNÇÃO QUE VERIFICA SE O CAMPO ESTÁ VAZIO
'** Recebe como parametro
'**** NOME DA TELA
'**** VETOR DE NOME dos campos
'**** VETOR DE DESCRIÇÃO dos campos
Public Function verificaCelulaVazia(nomeTela As String, listaDeCelulas() As Variant, listaDeDescricao() As Variant)
        
        '** Loop q percorre o vetor
        '** O UBound é onde acaba o vetor
        For contador = 0 To UBound(listaDeCelulas)
            
            '** Testa se o campo esta vazio
            '** IsEmpty verifica se a celula esta vazia ou não
            If IsEmpty(Sheets(nomeTela).Range(listaDeCelulas(contador)).Value) = True Then
                '** Mostra Mensagem na tela
                MsgBox "O campo " & listaDeDescricao(contador) & " está vazia"
                '** Cursor vai para o campo com erro
                Sheets(nomeTela).Range(listaDeCelulas(contador)).Select
                
                confirmaValidacao = False
                '** Sai da função e para a execução
                Exit Function
            End If
        
        Next contador
     
End Function



'** Função que verifica se o que foi digitado no campo é numero
'** Recebe como parametro
'**** NOME DA TELA
'**** VETOR que recebe a LISTA DE CAMPOS da tela
'**** VETOR que recebe a DESCRIÇÃO DOS CAMPOS da tela
'**** VETOR que recebe a LISTA DE CAMPOS que serão validados
Public Function verificaNumero(nomeTela As String, listaDeCelulas() As Variant, listaDeDescricao() As Variant, listaDeVerificacao() As Variant)

        '** Loop q percorre o vetor com o nome dos campos que serão verificados
        '** O UBound é onde acaba o vetor
        For ContadorVerificacao = 0 To UBound(listaDeVerificacao)
        
            '** Loop q percorre o vetor com o nome de todos os campos da tela
            '** O UBound é onde acaba o vetor
            For ContadorListaCelulas = 0 To UBound(listaDeCelulas)
                
                '** Testa se a informação digitada no campo é um NÚMERO
                '** If que encontra o campo a ser testado
                If (Sheets(nomeTela).Range(listaDeCelulas(ContadorListaCelulas)).Value) = (Sheets(nomeTela).Range(listaDeVerificacao(ContadorVerificacao)).Value) Then
                    
                    '** If que testa e o valor digitado no campo é um NÚMERO
                    If IsNumeric(Sheets(nomeTela).Range(listaDeCelulas(ContadorListaCelulas)).Value) = False Then
                        '** Mensagem caso o valor digitado no campo não for um NÚMERO
                        MsgBox "O campo " & listaDeDescricao(ContadorListaCelulas) & " só pode ser Número"
                        '** Cursor vai para o campo com erro
                        Sheets(nomeTela).Range(listaDeCelulas(ContadorListaCelulas)).Select
                        '** VARIAVEL que recebe a validação da verificação
                        confirmaValidacao = False
                        '** Sai da função e para a execução
                        Exit Function
                    End If
                        
                End If
            
            Next ContadorListaCelulas
        
        Next ContadorVerificacao
    
End Function
