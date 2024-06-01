Attribute VB_Name = "UnidadeMedida"

Sub CadastroUnidadeMedida()

    '** Cria e estancia um objeto tipo Validacao
    Dim validacao As New validacao
    '** Cria e estancia um objeto tipo cadastro
    Dim Cadastro As New Cadastro

    '** vetor que recebe o NOME dos campos
    Dim listaDeCelulas() As Variant
    '** vetor que recebe a DESCRIÇÃO dos campos
    Dim listaDeDescricao() As Variant

    '** Define o tamanho dos vetores
    ReDim listaDeCelulas(2)
    ReDim listaDeDescricao(2)
    
    '** Preenche o vetor de NOME dos campos
    listaDeCelulas(0) = "D4"
    listaDeCelulas(1) = "D6"
    listaDeCelulas(2) = "D8"
    
    '** Preenche o vetor de DESCRIÇÃO dos campos
    listaDeDescricao(0) = "ID"
    listaDeDescricao(1) = "DESCRIÇÃO"
    listaDeDescricao(2) = "ABREVIAÇÃO"
    
    '**Prepara para iniciar as validações
    validacao.iniciaValidacao
    
    '** Chama função que verifica se os campos estão vazios
    '** Passa como parametro
    '**** NOME DA TELA que contém os campos que serão verificados
    '**** VETOR que armazena o nome dos campos
    '**** VETOR que armazena a descrição dos campos
    validacao.verificaCelulaVazia "TelaCadastroUnidadeMedida", listaDeCelulas(), listaDeDescricao()
    
    '** Chama função que salva os dados na tabela
    '** Passa como parametro
    '**** NOME DA TELA que contém os campos com os dados que serão salvos
    '**** VETOR que armazena o nome os campos
    '**** NOME DA TABELA que armazena os dados
    '**** VAREAVEL DE VALIDAÇAO q verifica se todos os campos passaram ou não nos testes
    Cadastro.salvaDadosTabela "TelaCadastroUnidadeMedida", listaDeCelulas(), "TabelaCadastroUnidadeMedida", validacao.confirmaValidacao
    
    '** Salva a planilha
    ActiveWorkbook.Save
End Sub


Sub BuscarUnidadeMedida()

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
    
      '** Desprotege TelaBuscaUnidadeMedida
    ThisWorkbook.Sheets("TelaBuscaUnidadeMedida").Unprotect
    
    '** Limpa os campos da tela de busca
    '** Passa como parametros
    '**** NOME DA PRIMEIRA CELULA que recebera a busca
    '**** NOME DA TAELA que receberá a busca
    busca.limpaCelulasRecebeBusca "C8", "TelaBuscaUnidadeMedida"
    
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
    busca.buscaDados 2, "TelaBuscaUnidadeMedida", "TabelaCadastroUnidadeMedida", listaDeCelulas(), listaColunasTabela(), tipoBusca(), 8, 3, 5
    
    '** Protege TelaBuscaUnidadeMedida
    ThisWorkbook.Sheets("TelaBuscaUnidadeMedida").Protect
    
    '** Permite que as celulas bloqueadas sejam selecionadas
    ThisWorkbook.Sheets("TelaBuscaUnidadeMedida").EnableSelection = xlNoRestrictions
End Sub

Sub ApagarBuscaUnidadeMedida()

    
    '** Cria e estancia um objeto tipo Busca
    Dim busca As New busca
    '** Cria e estancia um objeto tipo funcionalidades
    Dim funcionalidades As New funcionalidades
    
    '** vetor que recebe o NOME dos campos
    Dim listaDeCelulas()
    '** Redimenciona o vetor
    ReDim listaDeCelulas(1)
    
    '** Desprotege TelaBuscaUnidadeMedida
    ThisWorkbook.Sheets("TelaBuscaUnidadeMedida").Unprotect
    
    '** Limpa os campos da tela de busca
    '** Passa como parametros
    '**** NOME DA PRIMEIRA CELULA que recebera a busca
    '**** NOME DA TAELA que receberá a busca
    busca.limpaCelulasRecebeBusca "C8", "TelaBuscaUnidadeMedida"
    
    '** Alimenta o vetor
    listaDeCelulas(0) = "C5"
    listaDeCelulas(1) = "D5"
    
    '** Apaga os campos de busca da tela
    '** Passa como parametros
    '**** NOME DA TELA de busca
    '**** VETOR que armazena o nome dos campos da tela de busca
    funcionalidades.apagaCampos "TelaBuscaUnidadeMedida", listaDeCelulas()
    
    '** Protege TelaBuscaUnidadeMedida
    ThisWorkbook.Sheets("TelaBuscaUnidadeMedida").Protect
    
End Sub

Sub EditarUnidadeMedida()

    '** Cria e estancia um objeto tipo edicao
    Dim edicao As New edicao
    
    '** Cria vetor que recebe os campos da TelaEditarUnidadeMedida
    Dim listaDeCelulasEditar() As Variant
    '** Redimenciona o vetor
    ReDim listaDeCelulasEditar(2)
    
     '** Desprotege TelaEditarUnidadeMedida
    ThisWorkbook.Sheets("TelaEditarUnidadeMedida").Unprotect

    '** Alimenta o vetor que recebe os campos da TelaEditarUnidadeMedida
    listaDeCelulasEditar(0) = "D4"
    listaDeCelulasEditar(1) = "D6"
    listaDeCelulasEditar(2) = "D8"
    
    '** Chama função da Classe edicao responsavel por encontrar e jogar na TELA DE EDIÇÃO o registro que será editado
    '** Passa como parametro
    '**** NUMERO DA COLUNA da tela de busca que recebe o primeiro registro
    '**** NOME DA TABELA que armazena os dados
    '**** NOME DA TELA de busca
    '**** NOME DA TELA que recebe os dados para edição
    '**** VETOR que armazena o nome das celulas da TELA DE EDIÇÃO
    edicao.encontraRegistro 3, "TabelaCadastroUnidadeMedida", "TelaBuscaUnidadeMedida", "TelaEditarUnidadeMedida", listaDeCelulasEditar()
    
    '** Protege TelaEditarMateriaPrima
    ThisWorkbook.Sheets("TelaEditarUnidadeMedida").Protect
    
End Sub

Sub CancelarEdicaoUnidadeMedida()

    '** Cria e estancia um objeto tipo funcionalidades
    Dim funcionalidades As New funcionalidades
    
    '** Cria vetor que recebe os campos da TelaEditarUnidadeMedida
    Dim listaDeCelulas() As Variant
    '** Redimenciona o vetor
    ReDim listaDeCelulas(2)
    
    '** Desprotege TelaEditarMateriaPrima
    ThisWorkbook.Sheets("TelaEditarUnidadeMedida").Unprotect
    
    '** Alimenta o vetor que recebe os campos da TelaEditarMateriaPrima
    listaDeCelulas(0) = "D4"
    listaDeCelulas(1) = "D6"
    listaDeCelulas(2) = "D8"
    
    '** Chama função da Classe Funcionalidades responsavel por apagar os dados dos campos
    '** Passa como parametro
    '**** NOME DA TELA que contem os campo
    '**** VETOR que armazena o nome dos campos
    funcionalidades.apagaCampos "TelaEditarUnidadeMedida", listaDeCelulas()
    
    '** Protege TelaEditarUnidadeMedida
    ThisWorkbook.Sheets("TelaEditarUnidadeMedida").Protect
    
    '** Permite que as celulas protegidas sejam selecionadas
    ThisWorkbook.Sheets("TelaBuscaUnidadeMedida").Activate
    
End Sub

Sub SalvaEdicaoUnidadeMedida()

    '** Cria e estancia um objeto tipo edicao
    Dim edicao As New edicao
     '** Cria e estancia um objeto tipo Validacao
    Dim validacao As New validacao
    '** Cria e estancia um objeto tipo funcionalidades
    Dim funcionalidades As New funcionalidades
    
    
    '** Cria vetor que recebe os campos da TelaEditarUnidadeMedida
    Dim listaDeCelulasEditar() As Variant
    '** Redimenciona o vetor
    ReDim listaDeCelulasEditar(2)
    '** Cria vetor que recebe a DESCRIÇÃO dos campos da TelaEditarUnidadeMedida
    Dim listaDeDescricaoEditar() As Variant
    '** Redimenciona o vetor
    ReDim listaDeDescricaoEditar(2)
    
    '** Desprotege a TelaEditarUnidadeMedida
    ThisWorkbook.Sheets("TelaEditarUnidadeMedida").Unprotect
    
    '** Preenche o vetor que recebe os campos da TelaEditarUnidadeMedida
    listaDeCelulasEditar(0) = "D4"
    listaDeCelulasEditar(1) = "D6"
    listaDeCelulasEditar(2) = "D8"
    
    '** Preenche o vetor que recebe a DESCRIÇÃO dos campos da TelaEditarUnidadeMedida
    listaDeDescricaoEditar(0) = "ID"
    listaDeDescricaoEditar(1) = "DESCRIÇÃO"
    listaDeDescricaoEditar(2) = "ABREVIAÇÃO"
    
    '**Prepara para iniciar as validações
    validacao.iniciaValidacao
    
    '** Chama função da Classe Edicao responsavel por salvar os dados alterados na TABELA
    '** Passa como parametro
    '**** NOME DA TELA de edição
    '**** VETOR com os campos da TELA DE EDIÇÃO
    '**** NOME DA TABELA
    '**** VARIAVEL que armazena se os campos passaram ou não pelas VALIDAÇÕES
    edicao.Editar "TelaEditarUnidadeMedida", listaDeCelulasEditar(), "TabelaCadastroUnidadeMedida", validacao.confirmaValidacao
    
    '** Apaga os campos da tela TelaEditarUnidadeMedida
    '** Passa como parametros
    '**** NOME DA TELA que contem os campos
    '**** VETOR que armazena o nome dos campos
    funcionalidades.apagaCampos "TelaEditarUnidadeMedida", listaDeCelulasEditar()
    
    '** Protege TelaEditarUnidadeMedida
    ThisWorkbook.Sheets("TelaEditarUnidadeMedida").Protect
    
    '** Chama TelaBuscaUnidadeMedida
    ThisWorkbook.Sheets("TelaBuscaUnidadeMedida").Activate
    
    '** Salva a planilha
    ActiveWorkbook.Save
    
    '** Chama função que apaga busca da tela TelaBuscaMateriaPrima
    Call ApagarBuscaUnidadeMedida
    
End Sub

Sub ChamaBuscaUnidadeMedida()
    
    '** Chama TelaBuscaUnidadeMedida
    ThisWorkbook.Sheets("TelaBuscaUnidadeMedida").Activate
    
    '** Chama função que apaga busca da tela ApagarBuscaUnidadeMedida
    Call ApagarBuscaUnidadeMedida
    
End Sub

