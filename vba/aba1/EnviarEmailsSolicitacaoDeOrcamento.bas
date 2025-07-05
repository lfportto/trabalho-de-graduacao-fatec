' ============================================================
' Nome do Projeto: Gestão no Processo de Solicitação de Compras, Recebimento e Estoque
' Autores: Luis Felipe Porto e Rodrigo da Silva Oliveira
' Instituição: Faculdade de Tecnologia de São José dos Campos - Prof. Jessen Vidal (FATEC SJC)
' Curso: Tecnologia em Gestão da Produção Industrial – 6º Semestre
' Descrição: Esta macro percorre a tabela da aba "Solicitação de orçamento" e, para cada item 
' com o status "Solicitar orçamento", gera um link personalizado para um formulário do Google Forms 
' com os campos Nome, Marca/Fornecedor e Quantidade já preenchidos. Em seguida, envia um e-mail 
' automático via Outlook para o destinatário informado, contendo o link do formulário e instruções 
' para preenchimento do orçamento. Após o envio, a macro atualiza o status do item para 
' "Aguardando orçamento" e registra a data/hora do envio.
' ============================================================

Sub EnviarEmailsSolicitacaoDeOrcamento()
    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim linha As ListRow
    Dim nomeItem As String, marca As String, quantidade As String, status As String
    Dim assunto As String, corpo As String
    Dim destinatario As String
    Dim encontrouItem As Boolean
    Dim linkForms As String, baseURL As String
    
    encontrouItem = False
    
    ' Inicializa o Outlook
    On Error Resume Next
    Set OutlookApp = GetObject(Class:="Outlook.Application")
    If OutlookApp Is Nothing Then
        Set OutlookApp = CreateObject(Class:="Outlook.Application")
    End If
    On Error GoTo 0
    
    If OutlookApp Is Nothing Then
        MsgBox "O Outlook não pôde ser iniciado.", vbExclamation, "Erro ao abrir o Outlook"
        Exit Sub
    End If
    
    Set ws = ThisWorkbook.Sheets("Solicitação de orçamento")
    Set tbl = ws.ListObjects("Table1")
    
    destinatario = "luis.porto01@fatec.sp.gov.br"
    
    ' URL base do Google Forms
    baseURL = "https://docs.google.com/forms/d/e/1FAIpQLSfxQcbM0Y5C6NhotQiO31GzJj_CQtIPTwGTboOwvvCkU5ztnQ/viewform?usp=pp_url"

    For Each linha In tbl.ListRows
        status = linha.Range(1, 4).Value
        
        If LCase(Trim(status)) = "solicitar orçamento" Then
            encontrouItem = True
            
            nomeItem = linha.Range(1, 1).Value ' Coluna C
            marca = linha.Range(1, 2).Value    ' Coluna D
            quantidade = linha.Range(1, 3).Value ' Coluna E
            
            ' Monta o link dinâmico do formulário com os dados codificados
            linkForms = baseURL & _
                "&entry.887929895=" & URLEncode(nomeItem) & _
                "&entry.329687793=" & URLEncode(marca) & _
                "&entry.217443874=" & URLEncode(quantidade)
            
            ' Monta o assunto e o corpo do e-mail
            assunto = "Solicitação de orçamento – " & nomeItem & " – " & marca
            corpo = "<p>Prezados,</p>" & _
            "<p>Gostaríamos de solicitar um orçamento para o(s) item(ns) abaixo:</p>" & _
            "<ul>" & _
            "<li><b>Item:</b> " & nomeItem & "</li>" & _
            "<li><b>Marca:</b> " & marca & "</li>" & _
            "<li><b>Quantidade:</b> " & quantidade & "</li>" & _
            "</ul>" & _
            "<p>Solicitamos, por gentileza, o preenchimento do formulário a seguir com as seguintes informações:</p>" & _
            "<ul>" & _
            "<li>Valor unitário</li>" & _
            "<li>Valor total</li>" & _
            "<li>Prazo de entrega</li>" & _
            "</ul>" & _
            "<p>Para facilitar o envio dessas informações, disponibilizamos o formulário a seguir:</p>" & _
            "<p><a href='" & linkForms & "'>Clique aqui para acessar o formulário</a></p>" & _
            "<p>Agradecemos desde já pela atenção e aguardamos o retorno com as informações solicitadas.</p>" & _
            "<p><b style='font-size: 11pt;'>PHS AUTOMAÇÃO E INSTALAÇÕES INDUSTRIAIS LTDA</b><br>" & _
            "<span style='font-size: 10pt;'>Rodrigo Oliveira<br>Técnico em Segurança do Trabalho</span></p>"

            Set OutlookMail = OutlookApp.CreateItem(0)
            With OutlookMail
                .To = destinatario
                .Subject = assunto
                .HTMLBody = corpo
                .Send
            End With
            
            linha.Range(1, 4).Value = "Aguardando orçamento"
            linha.Range(1, 5).Value = Now
        End If
    Next linha
    
    If encontrouItem Then
        MsgBox "E-mail(s) enviado(s) com sucesso!", vbInformation
    Else
        MsgBox "Nenhum item com status 'Solicitar orçamento'.", vbInformation
    End If
End Sub

Function URLEncode(str As String) As String
    URLEncode = Replace(str, " ", "+")
End Function
