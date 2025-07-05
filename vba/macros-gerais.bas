' ============================================================
' Nome do Projeto: Gestão no Processo de Solicitação de Compras, Recebimento e Estoque
' Autores: Luis Felipe Porto e Rodrigo da Silva Oliveira
' Instituição: Faculdade de Tecnologia de São José dos Campos - Prof. Jessen Vidal (FATEC SJC)
' Curso: Tecnologia em Gestão da Produção Industrial – 6º Semestre
' Descrição: Esta macro permite apagar as últimas linhas de uma tabela presente 
' na planilha ativa, mediante confirmação do usuário. A macro é genérica e funciona 
' em qualquer aba que contenha ao menos uma ListObject (tabela nomeada). 
' A quantidade de linhas a ser deletada é informada pelo usuário através de uma inputbox.
' ============================================================

Sub ApagarLinhasFinais()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim numLinhasApagar As Variant
    Dim totalLinhas As Long
    Dim i As Long
    
    ' Define a planilha ativa
    Set ws = ActiveSheet

    ' Verifica se há pelo menos uma tabela na planilha
    If ws.ListObjects.Count = 0 Then
        MsgBox "A planilha ativa não contém nenhuma tabela.", vbExclamation, "Tabela não encontrada"
        Exit Sub
    End If

    ' Usa a primeira tabela encontrada na planilha ativa
    Set tbl = ws.ListObjects(1)
    
    ' Verifica se a tabela tem dados
    totalLinhas = tbl.ListRows.Count
    If totalLinhas = 0 Then
        MsgBox "A tabela está vazia.", vbExclamation, "Nada a apagar"
        Exit Sub
    End If
    
    ' Solicita ao usuário o número de linhas a apagar
    numLinhasApagar = InputBox("Digite o número de linhas a apagar no final da tabela:", "Apagar Linhas")
    
    ' Cancelado ou em branco
    If numLinhasApagar = "" Then Exit Sub
    
    ' Verifica se é número válido
    If Not IsNumeric(numLinhasApagar) Or numLinhasApagar < 1 Or numLinhasApagar > totalLinhas Then
        MsgBox "Número de linhas a apagar é inválido ou maior que o total disponível (" & totalLinhas & ").", vbExclamation, "Erro"
        Exit Sub
    End If
    
    ' Confirmação
    If MsgBox("Deseja realmente apagar as últimas " & numLinhasApagar & " linha(s) da tabela '" & tbl.Name & "'?", vbYesNo + vbQuestion, "Confirmar exclusão") = vbNo Then Exit Sub

    ' Apaga de baixo para cima
    For i = 1 To numLinhasApagar
        tbl.ListRows(tbl.ListRows.Count).Delete
    Next i

    MsgBox numLinhasApagar & " linha(s) apagada(s) com sucesso!", vbInformation, "Concluído"
End Sub
