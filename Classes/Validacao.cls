VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Validacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'** IMPORTANTE PARA UTILIZA��O DA CLASSE **'
'**** Toda vez que uma fun��o recebe um vetor como parametro
'**** � necessario criar, redimencionar e alimentar esses vetores no local onde a fun��o ser� chamada
'**** para depois sim pass�-los como parametro


'** Variavel global da Classe que controla se passou pela valida��o
Public confirmaValidacao As Boolean

'** FUN��O CRIADA PARA INICIAR A VARIAVEL DE VALIDA��ES
Public Function iniciaValidacao()
    '** Variavel que controla se passou pela valida��o
    '** Inicia a variavel
    confirmaValidacao = True
End Function


'** FUN��O QUE VERIFICA SE O CAMPO EST� VAZIO
'** Recebe como parametro
'**** NOME DA TELA
'**** VETOR DE NOME dos campos
'**** VETOR DE DESCRI��O dos campos
Public Function verificaCelulaVazia(nomeTela As String, listaDeCelulas() As Variant, listaDeDescricao() As Variant)
        
        '** Loop q percorre o vetor
        '** O UBound � onde acaba o vetor
        For contador = 0 To UBound(listaDeCelulas)
            
            '** Testa se o campo esta vazio
            '** IsEmpty verifica se a celula esta vazia ou n�o
            If IsEmpty(Sheets(nomeTela).Range(listaDeCelulas(contador)).Value) = True Then
                '** Mostra Mensagem na tela
                MsgBox "O campo " & listaDeDescricao(contador) & " est� vazia"
                '** Cursor vai para o campo com erro
                Sheets(nomeTela).Range(listaDeCelulas(contador)).Select
                
                confirmaValidacao = False
                '** Sai da fun��o e para a execu��o
                Exit Function
            End If
        
        Next contador
     
End Function



'** Fun��o que verifica se o que foi digitado no campo � numero
'** Recebe como parametro
'**** NOME DA TELA
'**** VETOR que recebe a LISTA DE CAMPOS da tela
'**** VETOR que recebe a DESCRI��O DOS CAMPOS da tela
'**** VETOR que recebe a LISTA DE CAMPOS que ser�o validados
Public Function verificaNumero(nomeTela As String, listaDeCelulas() As Variant, listaDeDescricao() As Variant, listaDeVerificacao() As Variant)

        '** Loop q percorre o vetor com o nome dos campos que ser�o verificados
        '** O UBound � onde acaba o vetor
        For ContadorVerificacao = 0 To UBound(listaDeVerificacao)
        
            '** Loop q percorre o vetor com o nome de todos os campos da tela
            '** O UBound � onde acaba o vetor
            For ContadorListaCelulas = 0 To UBound(listaDeCelulas)
                
                '** Testa se a informa��o digitada no campo � um N�MERO
                '** If que encontra o campo a ser testado
                If (Sheets(nomeTela).Range(listaDeCelulas(ContadorListaCelulas)).Value) = (Sheets(nomeTela).Range(listaDeVerificacao(ContadorVerificacao)).Value) Then
                    
                    '** If que testa e o valor digitado no campo � um N�MERO
                    If IsNumeric(Sheets(nomeTela).Range(listaDeCelulas(ContadorListaCelulas)).Value) = False Then
                        '** Mensagem caso o valor digitado no campo n�o for um N�MERO
                        MsgBox "O campo " & listaDeDescricao(ContadorListaCelulas) & " s� pode ser N�mero"
                        '** Cursor vai para o campo com erro
                        Sheets(nomeTela).Range(listaDeCelulas(ContadorListaCelulas)).Select
                        '** VARIAVEL que recebe a valida��o da verifica��o
                        confirmaValidacao = False
                        '** Sai da fun��o e para a execu��o
                        Exit Function
                    End If
                        
                End If
            
            Next ContadorListaCelulas
        
        Next ContadorVerificacao
    
End Function
