Attribute VB_Name = "MateriaPrima"
Sub CadastrarMateriaPrima()
Attribute CadastrarMateriaPrima.VB_ProcData.VB_Invoke_Func = " \n14"
    
    '** Cria e estancia um objeto tipo Validacao
    Dim validacao As New validacao
    '** Cria e estancia um objeto tipo cadastro
    Dim Cadastro As New Cadastro
    
    '** vetor que recebe o NOME dos campos
    Dim listaDeCelulas() As Variant
    '** vetor que recebe a DESCRIÇÃO dos campos
    Dim listaDeDescricao() As Variant
    '** vetor que recebe o NOME dos campos expecificos que serão verificados
    Dim listaDeVerificacao() As Variant
    
    '** Define o tamanho dos vetores
    ReDim listaDeCelulas(4)
    ReDim listaDeDescricao(4)
    ReDim listaDeVerificacao(1)
    
    '** Preenche o vetor de NOME dos campos
    listaDeCelulas(0) = "D4"
    listaDeCelulas(1) = "D6"
    listaDeCelulas(2) = "D8"
    listaDeCelulas(3) = "D10"
    listaDeCelulas(4) = "D12"

    '** Preenche o vetor de DESCRIÇÃO dos campos
    listaDeDescricao(0) = "ID"
    listaDeDescricao(1) = "DESCRIÇÃO"
    listaDeDescricao(2) = "QUANTIDADE"
    listaDeDescricao(3) = "UNIDADE"
    listaDeDescricao(4) = "VALOR"
    
    '** Preenche o vetor de NOME DOS CAMPOS que serão verificados especificamente
    listaDeVerificacao(0) = "D8"
    listaDeVerificacao(1) = "D12"
    
    '**Prepara para iniciar as validações
    validacao.iniciaValidacao
    
    '** Chama função que verifica se os campos estão vazios
    '** Passa como parametro
    '**** NOME DA TELA que contém os campos que serão verificados
    '**** VETOR que armazena o nome dos campos
    '**** VETOR que armazena a descrição dos campos
    validacao.verificaCelulaVazia "TelaCadastroMateriaPrima", listaDeCelulas(), listaDeDescricao()
     

    '** Chama função que verifica se o conteúdo do campo é numero
    '** Passa como parametro
    '**** NOME DA TELA que contém os campos que serão verificados
    '**** VETOR que armazena o nome dos campos
    '**** VETOR que armazena a descrição dos campos
    '**** VETOR que armazena o nome dos campos que serão verificados
    validacao.verificaNumero "TelaCadastroMateriaPrima", listaDeCelulas(), listaDeDescricao(), listaDeVerificacao()
    
    '** Chama função que salva os dados na tabela
    '** Passa como parametro
    '**** NOME DA TELA que contém os campos com os dados que serão salvos
    '**** VETOR que armazena o nome os campos
    '**** NOME DA TABELA que armazena os dados
    '**** VAREAVEL DE VALIDAÇAO q verifica se todos os campos passaram ou não nos testes
    Cadastro.salvaDadosTabela "TelaCadastroMateriaPrima", listaDeCelulas(), "TabelaCadastroMateriaPrima", validacao.confirmaValidacao
    
    '** Salva a planilha
    ActiveWorkbook.Save
    
End Sub

Sub BuscarMateriaPrima()
Attribute BuscarMateriaPrima.VB_ProcData.VB_Invoke_Func = " \n14"
 
    '** Cria e estancia um objeto tipo Busca
    Dim busca As New busca
    
    '** Vetor de celulas da tela
    Dim listaDeCelulas() As Variant
    '** Vetor de colunas da tabela
    Dim listaColunasTabela() As Variant
    '** Vetor de tipo de busca
    Dim tipoBusca() As Integer
    
    '** Redimensiona os vetores
    ReDim listaDeCelulas(1)
    ReDim listaColunasTabela(1)
    ReDim tipoBusca(1)
    
    '** Desprotege TelaBuscaMateriaPrima
    ThisWorkbook.Sheets("TelaBuscaMateriaPrima").Unprotect
    
    '** Limpa os campos da tela de busca
    '** Passa como parametros
    '**** NOME DA PRIMEIRA CELULA que recebera a busca
    '**** NOME DA TAELA que receberá a busca
    busca.limpaCelulasRecebeBusca "C8", "TelaBuscaMateriaPrima"
        
    '** Alimenta os vetores
    listaDeCelulas(0) = "C5"
    listaColunasTabela(0) = 1
    tipoBusca(0) = 0
    
    listaDeCelulas(1) = "D5"
    listaColunasTabela(1) = 2
    tipoBusca(1) = 1
       
    '** Chama a função que realiza a busca
    '** Passa como parametros
    '**** NUMERO DE CAMPOS da tela de busca
    '**** NOME DA TELA de busca
    '**** NOME DA TABELA que contem os dados que serão verificados
    '**** VETOR que armazena o nome dos campos da tela de busca
    '**** VETOR que armazena a lista de colunas da tabela que contem os dados que serão verificados
    '**** VETOR que armazena o tipo de busca -- 0 para busca exata e 1 para busca de todos os resultados que contem a informação
    '**** NUMERO DA LINHA da tela que recebe a busca
    '**** NUMERO DA COLUNA da tela que recebe a busca
    '**** NUMERO DE CAMPOS da tabela
    busca.buscaDados 2, "TelaBuscaMateriaPrima", "TabelaCadastroMateriaPrima", listaDeCelulas(), listaColunasTabela(), tipoBusca(), 8, 3, 5
    
    '** Protege TelaBuscaMateriaPrima
    ThisWorkbook.Sheets("TelaBuscaMateriaPrima").Protect
    
    '** Permite que as celulas bloqueadas sejam selecionadas
    ThisWorkbook.Sheets("TelaBuscaMateriaPrima").EnableSelection = xlNoRestrictions
End Sub

Sub ApagarBuscaMateriaPrima()
Attribute ApagarBuscaMateriaPrima.VB_ProcData.VB_Invoke_Func = " \n14"

    '** Cria e estancia um objeto tipo Busca
    Dim busca As New busca
    '** Cria e estancia um objeto tipo funcionalidades
    Dim funcionalidades As New funcionalidades
    
    '** vetor que recebe o NOME dos campos
    Dim listaDeCelulas()
    '** Redimenciona o vetor
    ReDim listaDeCelulas(1)
    
    '** Desprotege TelaBuscaMateriaPrima
    ThisWorkbook.Sheets("TelaBuscaMateriaPrima").Unprotect
    
    '** Limpa os campos da tela de busca
    '** Passa como parametros
    '**** NOME DA PRIMEIRA CELULA que recebera a busca
    '**** NOME DA TAELA que receberá a busca
    busca.limpaCelulasRecebeBusca "C8", "TelaBuscaMateriaPrima"
    
    '** Alimenta o vetor
    listaDeCelulas(0) = "C5"
    listaDeCelulas(1) = "D5"
    
    '** Apaga os campos de busca da tela
    '** Passa como parametros
    '**** NOME DA TELA de busca
    '**** VETOR que armazena o nome dos campos da tela de busca
    funcionalidades.apagaCampos "TelaBuscaMateriaPrima", listaDeCelulas()
    
    '** Protege TelaBuscaMateriaPrima
    ThisWorkbook.Sheets("TelaBuscaMateriaPrima").Protect
    
End Sub


Sub EditarMateriaPrima()
    
    '** Cria e estancia um objeto tipo edicao
    Dim edicao As New edicao
    '** Cria e estancia um objeto tipo funcionalidades
    Dim funcionalidades As New funcionalidades
       
    
    '** Cria vetor que recebe os campos da TelaEditarMateriaPrima
    Dim listaDeCelulasEditar() As Variant
    '** Redimenciona o vetor
    ReDim listaDeCelulasEditar(4)
    
    '** Desprotege TelaEditarMateriaPrima
    ThisWorkbook.Sheets("TelaEditarMateriaPrima").Unprotect
    
    '** Alimenta o vetor que recebe os campos da TelaEditarMateriaPrima
    listaDeCelulasEditar(0) = "D4"
    listaDeCelulasEditar(1) = "D6"
    listaDeCelulasEditar(2) = "D8"
    listaDeCelulasEditar(3) = "D10"
    listaDeCelulasEditar(4) = "D12"
      
    '** Chama função da Classe edicao responsavel por encontrar e jogar na TELA DE EDIÇÃO o registro que será editado
    '** Passa como parametro
    '**** NUMERO DA COLUNA da tela de busca que recebe o primeiro registro
    '**** NOME DA TABELA que armazena os dados
    '**** NOME DA TELA de busca
    '**** NOME DA TELA que recebe os dados para edição
    '**** VETOR que armazena o nome das celulas da TELA DE EDIÇÃO
    edicao.encontraRegistro 3, "TabelaCadastroMateriaPrima", "TelaBuscaMateriaPrima", "TelaEditarMateriaPrima", listaDeCelulasEditar()
    
    '** Chama função da Classe funcionalidades responsavel por criar a LISTBOX
    '** Passa como parametro
    '**** NOME DA SHEET DA TABELA que vai popular a LISTBOX
    '**** NOME DA TABELA que vai popular a LISTBOX
    '**** NOME DA TELA onde será criada a LISTBOX
    '**** NOME DA CELULA onde que será a LISTBOX
    '**** NUMERO DA COLUNA DA TABELA que será usada para popuar a LISTBOX
    funcionalidades.criaListBox "TabelaCadastroUnidadeMedida", "TabelaCadastroUnidadeMedida", "TelaEditarMateriaPrima", "D10", 3
    
    '** Protege TelaEditarMateriaPrima
    ThisWorkbook.Sheets("TelaEditarMateriaPrima").Protect
    
End Sub


Sub CancelarEdicaoMateriaPrima()
    
    '** Cria e estancia um objeto tipo funcionalidades
    Dim funcionalidades As New funcionalidades
    
    '** Cria vetor que recebe os campos da TelaEditarMateriaPrima
    Dim listaDeCelulas() As Variant
    '** Redimenciona o vetor
    ReDim listaDeCelulas(4)
    
    '** Desprotege TelaEditarMateriaPrima
    ThisWorkbook.Sheets("TelaEditarMateriaPrima").Unprotect
    
    '** Alimenta o vetor que recebe os campos da TelaEditarMateriaPrima
    listaDeCelulas(0) = "D4"
    listaDeCelulas(1) = "D6"
    listaDeCelulas(2) = "D8"
    listaDeCelulas(3) = "D10"
    listaDeCelulas(4) = "D12"
    
    '** Chama função da Classe Funcionalidades responsavel por apagar os dados dos campos
    '** Passa como parametro
    '**** NOME DA TELA que contem os campo
    '**** VETOR que armazena o nome dos campos
    funcionalidades.apagaCampos "TelaEditarMateriaPrima", listaDeCelulas()
    
    '** Protege TelaEditarMateriaPrima
    ThisWorkbook.Sheets("TelaEditarMateriaPrima").Protect
    
    '** Permite que as celulas protegidas sejam selecionadas
    ThisWorkbook.Sheets("TelaBuscaMateriaPrima").Activate

End Sub



Sub SalvaEdicaoMateriaPrima()
    
    '** Cria e estancia um objeto tipo edicao
    Dim edicao As New edicao
     '** Cria e estancia um objeto tipo Validacao
    Dim validacao As New validacao
    '** Cria e estancia um objeto tipo funcionalidades
    Dim funcionalidades As New funcionalidades
    
    '** Cria vetor que recebe os campos da TelaEditarMateriaPrima
    Dim listaDeCelulasEditar() As Variant
    '** Redimenciona o vetor
    ReDim listaDeCelulasEditar(4)
    '** Cria vetor que recebe a DESCRIÇÃO dos campos da TelaEditarMateriaPrima
    Dim listaDeDescricaoEditar() As Variant
    '** Redimenciona o vetor
    ReDim listaDeDescricaoEditar(4)
    '** Cria vetor que recebe o NOME dos campos expecificos que serão validados da TelaEditarMateriaPrima
    Dim listaDeVerificacao() As Variant
    '** Redimenciona o vetor
    ReDim listaDeVerificacao(1)
    
    '** Desprotege a TelaEditarMateriaPrima
    ThisWorkbook.Sheets("TelaEditarMateriaPrima").Unprotect
    
    '** Preenche o vetor que recebe os campos da TelaEditarMateriaPrima
    listaDeCelulasEditar(0) = "D4"
    listaDeCelulasEditar(1) = "D6"
    listaDeCelulasEditar(2) = "D8"
    listaDeCelulasEditar(3) = "D10"
    listaDeCelulasEditar(4) = "D12"
    
    '** Preenche o vetor que recebe a DESCRIÇÃO dos campos da TelaEditarMateriaPrima
    listaDeDescricaoEditar(0) = "ID"
    listaDeDescricaoEditar(1) = "DESCRIÇÃO"
    listaDeDescricaoEditar(2) = "QUANTIDADE"
    listaDeDescricaoEditar(3) = "UNIDADE"
    listaDeDescricaoEditar(4) = "VALOR"

    '** Preenche o vetor de NOME DOS CAMPOS que serão validados
    listaDeVerificacao(0) = "D8"
    listaDeVerificacao(1) = "D12"
    
    
    '**Prepara para iniciar as validações
    validacao.iniciaValidacao
    
    '** Chama função que verifica se os campos estão vazios
    '** Passa como parametro
    '**** NOME DA TELA que contém os campos que serão verificados
    '**** VETOR que armazena o nome dos campos
    '**** VETOR que armazena a descrição dos campos
    validacao.verificaCelulaVazia "TelaEditarMateriaPrima", listaDeCelulasEditar(), listaDeDescricaoEditar()

    
    '** Chama função que verifica se o conteúdo do campo é numero
    '** Passa como parametro
    '**** NOME DA TELA que contém os campos que serão verificados
    '**** VETOR que armazena o nome dos campos
    '**** VETOR que armazena a descrição dos campos
    '**** VETOR que armazena o nome dos campos que serão verificados
    validacao.verificaNumero "TelaEditarMateriaPrima", listaDeCelulasEditar(), listaDeDescricaoEditar(), listaDeVerificacao()
    
    '** Chama função da Classe Edicao responsavel por salvar os dados alterados na TABELA
    '** Passa como parametro
    '**** NOME DA TELA de edição
    '**** VETOR com os campos da TELA DE EDIÇÃO
    '**** NOME DA TABELA
    '**** VARIAVEL que armazena se os campos passaram ou não pelas VALIDAÇÕES
    edicao.Editar "TelaEditarMateriaPrima", listaDeCelulasEditar(), "TabelaCadastroMateriaPrima", validacao.confirmaValidacao
    
    '** Apaga os campos da tela TelaEditarMateriaPrima
    '** Passa como parametros
    '**** NOME DA TELA que contem os campos
    '**** VETOR que armazena o nome dos campos
    funcionalidades.apagaCampos "TelaEditarMateriaPrima", listaDeCelulasEditar()
    
    '** Protege TelaEditarMateriaPrima
    ThisWorkbook.Sheets("TelaEditarMateriaPrima").Protect
    
    '** Chama TelaBuscaMateriaPrima
    ThisWorkbook.Sheets("TelaBuscaMateriaPrima").Activate
    
    '** Salva a planilha
    ActiveWorkbook.Save
    
    '** Chama função que apaga busca da tela TelaBuscaMateriaPrima
    Call ApagarBuscaMateriaPrima
End Sub


Sub ChamaBuscaMateriaPrima()
    
    '** Chama TelaBuscaMateriaPrima
    ThisWorkbook.Sheets("TelaBuscaMateriaPrima").Activate
    
    '** Chama função que apaga busca da tela TelaBuscaMateriaPrima
    Call ApagarBuscaMateriaPrima
    
End Sub

Sub ChamaCadastroMateriaPrima()
    
    '** Cria e estancia um objeto tipo funcionalidades
    Dim funcionalidades As New funcionalidades
       
    '** Chama função da Classe funcionalidades responsavel por criar a LISTBOX
    '** Passa como parametro
    '**** NOME DA SHEET DA TABELA que vai popular a LISTBOX
    '**** NOME DA TABELA que vai popular a LISTBOX
    '**** NOME DA TELA onde será criada a LISTBOX
    '**** NOME DA CELULA onde que será a LISTBOX
    '**** NUMERO DA COLUNA DA TABELA que será usada para popuar a LISTBOX
    funcionalidades.criaListBox "TabelaCadastroUnidadeMedida", "TabelaCadastroUnidadeMedida", "TelaCadastroMateriaPrima", "D10", 3
    
    '** Chama TelaBuscaMateriaPrima
    ThisWorkbook.Sheets("TelaCadastroMateriaPrima").Activate
    
End Sub


