' ============================================================
' Nome do Projeto: Gestão no Processo de Solicitação de Compras, Recebimento e Estoque
' Autores: Luis Felipe Porto e Rodrigo da Silva Oliveira
' Instituição: Faculdade de Tecnologia de São José dos Campos - Prof. Jessen Vidal (FATEC SJC)
' Curso: Tecnologia em Gestão da Produção Industrial – 6º Semestre
' Descrição:
' ============================================================

Sub MarcarPedidoComoRecebido()

    Dim ws As Worksheet
    Dim ticketIDBusca As String
    Dim ticketIDPlanilha As String
    Dim ultimaLinha As Long
    Dim i As Long
    Dim encontrado As Boolean

    Set ws = ThisWorkbook.Sheets("Pedidos aprovados")
    
    ' Solicita o ID com validação de entrada
    ticketIDBusca = InputBox("Digite o Ticket ID do pedido que foi recebido:", "Marcar como Recebido")
    
    If ticketIDBusca = "" Then Exit Sub ' Cancelado ou vazio
    
    ' Limpa espaços, aspas, e caracteres invisíveis como Chr(160) (espaço não separável)
    ticketIDBusca = Trim(Replace(ticketIDBusca, Chr(160), ""))
    ticketIDBusca = Replace(ticketIDBusca, Chr(10), "") ' quebra de linha
    ticketIDBusca = Replace(ticketIDBusca, Chr(13), "") ' retorno de carro
    ticketIDBusca = Replace(ticketIDBusca, "'", "")     ' apóstrofo (se o usuário colar com apóstrofo)
    
    ultimaLinha = ws.Cells(ws.Rows.Count, "H").End(xlUp).Row
    encontrado = False
    
    ' Garante que a coluna H está como texto (evita comparação inconsistente)
    ws.Columns("H").NumberFormat = "@"
    
    For i = 8 To ultimaLinha ' Começa na linha 8, após o cabeçalho

        ' Garante leitura como texto + limpa qualquer ruído invisível
        ticketIDPlanilha = Trim(Replace(CStr(ws.Cells(i, "H").Value), Chr(160), ""))
        ticketIDPlanilha = Replace(ticketIDPlanilha, Chr(10), "")
        ticketIDPlanilha = Replace(ticketIDPlanilha, Chr(13), "")
        ticketIDPlanilha = Replace(ticketIDPlanilha, "'", "")
        
        If ticketIDPlanilha = ticketIDBusca Then
            ' Marca como recebido e insere data e hora
            ws.Cells(i, "F").Value = "Recebido"
            ws.Cells(i, "G").Value = Format(Now, "dd/mm/yy HH:mm")
            encontrado = True
            Exit For
        End If
    Next i

    ' Mensagem de feedback
    If encontrado Then
        MsgBox "Pedido com Ticket ID '" & ticketIDBusca & "' marcado como RECEBIDO com sucesso!", vbInformation, "Atualização concluída"
    Else
        MsgBox "Nenhum pedido encontrado com o Ticket ID informado.", vbExclamation, "ID não encontrado"
    End If

End Sub
