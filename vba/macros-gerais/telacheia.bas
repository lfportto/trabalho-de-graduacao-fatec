' ============================================================
' Nome do Projeto: Gestão no Processo de Solicitação de Compras, Recebimento e Estoque
' Autores: Luis Felipe Porto e Rodrigo da Silva Oliveira
' Instituição: Faculdade de Tecnologia de São José dos Campos - Prof. Jessen Vidal (FATEC SJC)
' Curso: Tecnologia em Gestão da Produção Industrial – 6º Semestre
' Descrição: Esta macro ajusta a visualização do Excel para o modo tela cheia,
' ocultando a barra de fórmulas e as guias das planilhas. Ideal para apresentações,
' dashboards ou uso em ambientes controlados onde se deseja focar apenas no conteúdo.
' ============================================================

Sub telacheia()
    Application.DisplayFullScreen = 1
    Application.DisplayFormulaBar = 0
    ActiveWindow.DisplayWorkbookTabs = 0
End Sub
