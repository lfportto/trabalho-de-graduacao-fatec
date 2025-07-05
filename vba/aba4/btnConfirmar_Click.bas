' ============================================================
' Nome do Projeto: Gestão no Processo de Solicitação de Compras, Recebimento e Estoque
' Autores: Luis Felipe Porto e Rodrigo da Silva Oliveira
' Instituição: Faculdade de Tecnologia de São José dos Campos - Prof. Jessen Vidal (FATEC SJC)
' Curso: Tecnologia em Gestão da Produção Industrial – 6º Semestre
' Descrição:
' ============================================================

Private Sub btnConfirmar_Click()
    Dim wsEstoque As Worksheet
    Dim ultimaLinha As Long
    Dim ticketIDInput As String
    Dim ticketIDPlanilha As String
    Dim quantidadeInput As Double
    Dim solicitanteInput As String
    Dim nomeItem As String
    Dim marcaFornecedor As String
    Dim linhaFonte As Long
    Dim i As Long
    Dim encontrado As Boolean
    Dim saldoDisponivel As Double

    Set wsEstoque = ThisWorkbook.Sheets("Estoque")

    ' Captura os valores do formulário
    ticketIDInput = Trim(Me.txtTicketID.Value)
    solicitanteInput = Trim(Me.txtSolicitante.Value)
    quantidadeInput = Val(Me.txtQuantidade.Value)

    ' Limpeza do Ticket ID
    ticketIDInput = Replace(ticketIDInput, Chr(160), "")
    ticketIDInput = Replace(ticketIDInput, Chr(10), "")
    ticketIDInput = Replace(ticketIDInput, Chr(13), "")
    ticketIDInput = Replace(ticketIDInput, "'", "")

    If ticketIDInput = "" Or solicitanteInput = "" Or quantidadeInput <= 0 Then
        MsgBox "Por favor, preencha todos os campos corretamente.", vbExclamation
        Exit Sub
    End If

    ' Garante que a coluna H está como texto
    wsEstoque.Columns("H").NumberFormat = "@"

    ' Inicializa variáveis
    encontrado = False
    saldoDisponivel = 0

    ' Percorre todas as linhas do estoque para somar as quantidades do Ticket ID
    For i = 8 To wsEstoque.Cells(wsEstoque.Rows.Count, "H").End(xlUp).Row
        ticketIDPlanilha = Trim(CStr(wsEstoque.Cells(i, "H").Value))
        ticketIDPlanilha = Replace(ticketIDPlanilha, Chr(160), "")
        ticketIDPlanilha = Replace(ticketIDPlanilha, Chr(10), "")
        ticketIDPlanilha = Replace(ticketIDPlanilha, Chr(13), "")
        ticketIDPlanilha = Replace(ticketIDPlanilha, "'", "")

        If ticketIDPlanilha = ticketIDInput Then
            saldoDisponivel = saldoDisponivel + Val(wsEstoque.Cells(i, "E").Value)
            If Not encontrado Then
                nomeItem = wsEstoque.Cells(i, "C").Value
                marcaFornecedor = wsEstoque.Cells(i, "D").Value
                encontrado = True
            End If
        End If
    Next i

    If Not encontrado Then
        MsgBox "Ticket ID não encontrado no estoque.", vbExclamation
        Exit Sub
    End If

    ' Verifica se há saldo suficiente
    If quantidadeInput > saldoDisponivel Then
        MsgBox "Não é possível retirar " & quantidadeInput & " unidade(s). Saldo disponível para esse item: " & saldoDisponivel & ".", vbCritical, "Saldo insuficiente"
        Exit Sub
    End If

    ' Adiciona nova linha com a baixa
    ultimaLinha = wsEstoque.Cells(wsEstoque.Rows.Count, "C").End(xlUp).Row + 1

    wsEstoque.Cells(ultimaLinha, "C").Value = nomeItem
    wsEstoque.Cells(ultimaLinha, "D").Value = marcaFornecedor
    wsEstoque.Cells(ultimaLinha, "E").Value = -Abs(quantidadeInput)
    wsEstoque.Cells(ultimaLinha, "F").Value = solicitanteInput
    wsEstoque.Cells(ultimaLinha, "G").Value = Now
    wsEstoque.Cells(ultimaLinha, "G").NumberFormat = "dd/mm/yyyy hh:mm"
    wsEstoque.Cells(ultimaLinha, "H").NumberFormat = "@"
    wsEstoque.Cells(ultimaLinha, "H").Value = ticketIDInput

    MsgBox "Baixa no estoque registrada com sucesso!", vbInformation

    Unload Me
End Sub
