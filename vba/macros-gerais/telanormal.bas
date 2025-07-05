' ============================================================
' Nome do Projeto: Gestão no Processo de Solicitação de Compras, Recebimento e Estoque
' Autores: Luis Felipe Porto e Rodrigo da Silva Oliveira
' Instituição: Faculdade de Tecnologia de São José dos Campos - Prof. Jessen Vidal (FATEC SJC)
' Curso: Tecnologia em Gestão da Produção Industrial – 6º Semestre
' Descrição: Esta macro retorna à visualização padrão do Excel, desativando o modo tela cheia,
' exibindo novamente a barra de fórmulas e as guias das planilhas. Ideal para edição ou 
' navegação comum do usuário fora do modo de apresentação.
' ============================================================

Sub telanormal()
    Application.DisplayFullScreen = 0
    Application.DisplayFormulaBar = 1
    ActiveWindow.DisplayWorkbookTabs = 1
End Sub
