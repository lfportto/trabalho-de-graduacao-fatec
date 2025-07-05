' ============================================================
' Nome do Projeto: Gestão no Processo de Solicitação de Compras, Recebimento e Estoque
' Autores: Luis Felipe Porto e Rodrigo da Silva Oliveira
' Instituição: Faculdade de Tecnologia de São José dos Campos - Prof. Jessen Vidal (FATEC SJC)
' Curso: Tecnologia em Gestão da Produção Industrial – 6º Semestre
' Descrição:
' ============================================================

Sub TransferirParaEstoque()

    Dim wsOrigem As Worksheet
    Dim wsDestino As Worksheet
    Dim ultimaLinhaOrigem As Long, ultimaLinhaDestino As Long
    Dim i As Long
    Dim ticketID As String
    Dim ticketIDsDestino As Object
    Dim registrosTransferidos As Long

    Set wsOrigem = ThisWorkbook.Sheets("Pedidos aprovados")
    Set wsDestino = ThisWorkbook.Sheets("Estoque")
    
    ultimaLinhaOrigem = wsOrigem.Cells(wsOrigem.Rows.Count, "C").End(xlUp).Row
    ultimaLinhaDestino = wsDestino.Cells(wsDestino.Rows.Count, "C").End(xlUp).Row

    Set ticketIDsDestino = CreateObject("Scripting.Dictionary")
    
    ' Carrega os Ticket IDs já existentes na aba Estoque (coluna H)
    For i = 8 To ultimaLinhaDestino
        ticketID = Trim(wsDestino.Cells(i, "H").Text)
        If ticketID <> "" Then
            ticketIDsDestino(ticketID) = True
        End If
    Next i

    registrosTransferidos = 0

    ' Verifica os registros com status "Recebido" na aba Pedidos aprovados
    For i = 8 To ultimaLinhaOrigem
        If Trim(wsOrigem.Cells(i, "F").Value) = "Recebido" Then
            ticketID = Trim(CStr(wsOrigem.Cells(i, "H").Value))
            
            If ticketID <> "" And Not ticketIDsDestino.exists(ticketID) Then
                ultimaLinhaDestino = ultimaLinhaDestino + 1
                
                ' Copia os dados para a aba Estoque
                wsDestino.Cells(ultimaLinhaDestino, "C").Value = wsOrigem.Cells(i, "C").Value ' Nome
                wsDestino.Cells(ultimaLinhaDestino, "D").Value = wsOrigem.Cells(i, "D").Value ' Marca / Fornecedor
                wsDestino.Cells(ultimaLinhaDestino, "E").Value = wsOrigem.Cells(i, "E").Value ' Quantidade
                wsDestino.Cells(ultimaLinhaDestino, "F").Value = "" ' Solicitante em branco
                wsDestino.Cells(ultimaLinhaDestino, "G").Value = wsOrigem.Cells(i, "G").Value ' Data da entrega
                wsDestino.Cells(ultimaLinhaDestino, "H").NumberFormat = "@" ' Ticket ID como texto
                wsDestino.Cells(ultimaLinhaDestino, "H").Value = CStr(ticketID)
                
                registrosTransferidos = registrosTransferidos + 1
            End If
        End If
    Next i

    If registrosTransferidos > 0 Then
        MsgBox registrosTransferidos & " item(ns) transferido(s) para o estoque com sucesso!", vbInformation, "Transferência concluída"
    Else
        MsgBox "Nenhum novo item com status 'Recebido' encontrado para transferir.", vbExclamation, "Sem transferências"
    End If

End Sub
