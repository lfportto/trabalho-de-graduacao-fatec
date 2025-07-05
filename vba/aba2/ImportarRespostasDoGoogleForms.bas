' ============================================================
' Nome do Projeto: Gestão no Processo de Solicitação de Compras, Recebimento e Estoque
' Autores: Luis Felipe Porto e Rodrigo da Silva Oliveira
' Instituição: Faculdade de Tecnologia de São José dos Campos - Prof. Jessen Vidal (FATEC SJC)
' Curso: Tecnologia em Gestão da Produção Industrial – 6º Semestre
' Descrição:
' ============================================================

Sub ImportarRespostasDoGoogleForms()
    Dim http As Object
    Dim csvConteudo As String
    Dim linhas() As String
    Dim colunas() As String
    Dim i As Long
    Dim wsOrcados As Worksheet
    Dim wsSolicitacao As Worksheet
    Dim ultimaLinha As Long
    Dim ticketIDAtual As String
    Dim celula As Range
    Dim respostaJaExiste As Boolean
    Dim novosRegistros As Long
    Dim nomeItem As String
    Dim marcaItem As String
    Dim linhaSolicitacao As Long
    Dim ultimaLinhaSolicitacao As Long

    Set wsOrcados = ThisWorkbook.Sheets("Itens orçados")
    Set wsSolicitacao = ThisWorkbook.Sheets("Solicitação de orçamento")

    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", "https://docs.google.com/spreadsheets/d/1dxcbUpPhCg1dP7Ij_fjqTcGX0jRdVX195T2fq2hnqGU/export?format=csv", False
    http.Send

    If http.status = 200 Then
        csvConteudo = http.responseText
        linhas = Split(csvConteudo, vbLf)

        novosRegistros = 0

        For i = 1 To UBound(linhas)
            If Trim(linhas(i)) <> "" Then
                colunas = Split(linhas(i), ",")

                ticketIDAtual = Trim(colunas(7)) ' Ticket ID
                nomeItem = Trim(colunas(1))      ' Nome do item
                marcaItem = Trim(colunas(2))     ' Marca / Fornecedor

                ' Verifica duplicidade
                respostaJaExiste = False
                For Each celula In wsOrcados.Range("K5:K" & wsOrcados.Cells(wsOrcados.Rows.Count, "K").End(xlUp).Row)
                    If Trim(celula.Value) = ticketIDAtual And ticketIDAtual <> "" Then
                        respostaJaExiste = True
                        Exit For
                    End If
                Next celula

                ' Se for novo registro:
                If Not respostaJaExiste And ticketIDAtual <> "" Then
                    ultimaLinha = wsOrcados.Cells(wsOrcados.Rows.Count, "C").End(xlUp).Row + 1

                    wsOrcados.Cells(ultimaLinha, "C").Value = nomeItem
                    wsOrcados.Cells(ultimaLinha, "D").Value = marcaItem
                    wsOrcados.Cells(ultimaLinha, "E").Value = colunas(3) ' Quantidade
                    wsOrcados.Cells(ultimaLinha, "F").Value = colunas(4) ' Valor unitário
                    wsOrcados.Cells(ultimaLinha, "H").Value = colunas(5) ' Prazo
                    wsOrcados.Cells(ultimaLinha, "J").Value = colunas(0) ' Data/hora
                    wsOrcados.Cells(ultimaLinha, "K").Value = "'" & ticketIDAtual

                    novosRegistros = novosRegistros + 1

                    ' Agora: atualiza o status na aba "Solicitação de orçamento"
                    ultimaLinhaSolicitacao = wsSolicitacao.Cells(wsSolicitacao.Rows.Count, "C").End(xlUp).Row
                    For linhaSolicitacao = 8 To ultimaLinhaSolicitacao
                        If Trim(wsSolicitacao.Cells(linhaSolicitacao, "C").Value) = nomeItem And _
                           Trim(wsSolicitacao.Cells(linhaSolicitacao, "D").Value) = marcaItem Then
                            wsSolicitacao.Cells(linhaSolicitacao, "F").Value = "Pedido orçado"
                            Exit For
                        End If
                    Next linhaSolicitacao
                End If
            End If
        Next i

        If novosRegistros > 0 Then
            MsgBox novosRegistros & " novo(s) registro(s) importado(s) com sucesso!", vbInformation, "Importação de itens"
        Else
            MsgBox "Todos os registros já foram importados.", vbInformation, "Nenhuma nova importação"
        End If
    Else
        MsgBox "Erro ao baixar respostas: " & http.status, vbCritical, "Erro de Download"
    End If
End Sub

' Agendamento automático
Sub IniciarAtualizacaoAutomatica()
    Dim proximaAtualizacao As Date
    proximaAtualizacao = Now + TimeValue("01:00:00") ' 1 hora
    Application.OnTime proximaAtualizacao, "ImportarRespostasDoGoogleForms"
End Sub

' Cancelar agendamento automático
Sub CancelarAtualizacaoAutomatica()
    On Error Resume Next
    Dim proximaAtualizacao As Date
    proximaAtualizacao = Now + TimeValue("01:00:00") ' Precisa indicar o mesmo horário agendado
    Application.OnTime proximaAtualizacao, "ImportarRespostasDoGoogleForms", , False
End Sub
