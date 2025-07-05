' ============================================================
' Nome do Projeto: Gestão no Processo de Solicitação de Compras, Recebimento e Estoque
' Autores: Luis Felipe Porto e Rodrigo da Silva Oliveira
' Instituição: Faculdade de Tecnologia de São José dos Campos - Prof. Jessen Vidal (FATEC SJC)
' Curso: Tecnologia em Gestão da Produção Industrial – 6º Semestre
' Descrição: Esta macro transfere automaticamente os itens aprovados na aba "Itens orçados"
' para a aba "Pedidos aprovados". A macro verifica se o status do pedido é "Aprovado" e
' se ainda não foi transferido, evitando duplicidade com base no Ticket ID.
' Ao transferir, marca o item como "Não recebido" e registra que já foi processado.
' ============================================================

Sub TransferirPedidosAprovados()
    Dim wsOrigem As Worksheet
    Dim wsDestino As Worksheet
    Dim ultimaLinhaOrigem As Long, ultimaLinhaDestino As Long
    Dim i As Long
    Dim ticketID As String

    Set wsOrigem = ThisWorkbook.Sheets("Itens orçados")
    Set wsDestino = ThisWorkbook.Sheets("Pedidos aprovados")
    
    ultimaLinhaOrigem = wsOrigem.Cells(wsOrigem.Rows.Count, "C").End(xlUp).Row
    ultimaLinhaDestino = wsDestino.Cells(wsDestino.Rows.Count, "C").End(xlUp).Row
    
    Dim ticketIDsDestino As Object
    Set ticketIDsDestino = CreateObject("Scripting.Dictionary")
    
    ' Carrega os Ticket IDs já existentes no destino (coluna H)
    For i = 8 To ultimaLinhaDestino
        ticketID = Trim(wsDestino.Cells(i, "H").Text)
        If ticketID <> "" Then
            ticketIDsDestino(ticketID) = True
        End If
    Next i
    
    Dim registrosTransferidos As Long
    registrosTransferidos = 0

    ' Verifica os registros aprovados na origem
    For i = 5 To ultimaLinhaOrigem ' Começa da linha 5 para pular o cabeçalho
        If wsOrigem.Cells(i, "I").Value = "Aprovado" And wsOrigem.Cells(i, "L").Value <> "Sim" Then
            ticketID = Trim(CStr(wsOrigem.Cells(i, "K").Value)) ' Coluna K da aba "Itens orçados" = Ticket ID
            
            If ticketID <> "" And Not ticketIDsDestino.exists(ticketID) Then
                ultimaLinhaDestino = ultimaLinhaDestino + 1
                
                ' Copia os dados nas colunas corretas
                wsDestino.Cells(ultimaLinhaDestino, "C").Value = wsOrigem.Cells(i, "C").Value ' Nome do item
                wsDestino.Cells(ultimaLinhaDestino, "D").Value = wsOrigem.Cells(i, "D").Value ' Marca/Fornecedor
                wsDestino.Cells(ultimaLinhaDestino, "E").Value = wsOrigem.Cells(i, "E").Value ' Quantidade
                wsDestino.Cells(ultimaLinhaDestino, "F").Value = "Não recebido" ' Status do pedido
                wsDestino.Cells(ultimaLinhaDestino, "G").Value = "" ' Data da entrega (vazio)

                ' Garante que o Ticket ID seja texto puro
                wsDestino.Cells(ultimaLinhaDestino, "H").NumberFormat = "@"
                wsDestino.Cells(ultimaLinhaDestino, "H").Value = CStr(ticketID)

                ' Marca como transferido na aba de origem
                wsOrigem.Cells(i, "L").Value = "Sim"
                
                registrosTransferidos = registrosTransferidos + 1
            End If
        End If
    Next i
    
    If registrosTransferidos > 0 Then
        MsgBox registrosTransferidos & " pedido(s) aprovado(s) transferido(s) com sucesso!", vbInformation, "Transferência concluída"
    Else
        MsgBox "Nenhum novo pedido aprovado para transferir.", vbExclamation, "Nada transferido"
    End If

End Sub
