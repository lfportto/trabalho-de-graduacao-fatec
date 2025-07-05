' ============================================================
' Nome do Projeto: Gestão no Processo de Solicitação de Compras, Recebimento e Estoque
' Autores: Luis Felipe Porto e Rodrigo da Silva Oliveira
' Instituição: Faculdade de Tecnologia de São José dos Campos - Prof. Jessen Vidal (FATEC SJC)
' Curso: Tecnologia em Gestão da Produção Industrial – 6º Semestre
' Descrição:
' ============================================================

Sub registrarNovoItem()
    Dim linhaDestino As Long
    Dim resposta As VbMsgBoxResult
    
    ' Verifica se algum dos campos obrigatórios está vazio (C5, D5, E5)
    If Range("C5") = "" Or Range("D5") = "" Or Range("E5") = "" Then
        resposta = MsgBox("Há campos não preenchidos. Deseja confirmar o registro mesmo assim?", vbYesNo + vbQuestion, "Confirmação de Registro")
        If resposta <> vbYes Then Exit Sub
    End If

    ' Encontra a próxima linha vazia na coluna C (assumindo que a tabela começa em C8)
    linhaDestino = Range("C8").End(xlDown).Row + 1

    ' Copia os valores preenchidos
    Cells(linhaDestino, 3).Value = Range("C5").Value ' Nome
    Cells(linhaDestino, 4).Value = Range("D5").Value ' Marca / Fornecedor
    Cells(linhaDestino, 5).Value = Range("E5").Value ' Quantidade
    Cells(linhaDestino, 6).Value = "Solicitar orçamento" ' Status do pedido

    ' Limpa os campos preenchidos
    Range("C5:E5").ClearContents

    ' Mensagem de sucesso
    MsgBox "Registro realizado com sucesso!", vbInformation, "Novo item registrado"
    
    ' Coloca o cursor de volta no primeiro campo
    Range("C5").Select
End Sub
