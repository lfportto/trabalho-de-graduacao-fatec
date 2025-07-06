# Sistema de Gestão de Solicitação de Compras, Recebimento e Estoque
Este repositório contém os códigos desenvolvidos em `VBA` (Visual Basic for Applications) e `Google Apps Script` como parte do Trabalho de Graduação apresentado à Faculdade de Tecnologia de São José dos Campos – Prof. Jessen Vidal (FATEC SJC), durante o 5º e 6º período do curso de Tecnologia em Gestão da Produção Industrial, pelos alunos Luis Felipe Porto e Rodrigo da Silva Oliveira.

O objetivo do projeto é automatizar e padronizar o processo de solicitação de compras, orçamentos, aprovação, recebimento e controle de estoque de materiais de uma empresa de porte pequeno, por meio de uma planilha integrada com formulários e um dashboard gerencial para monitoramento e interação no Power BI.

O fluxo completo é descrito no Trabalho de Graduação [disponível aqui](https://drive.google.com/file/d/1il2iBtzbF8Q_8AwS4Z1RSmsinUmUGz2x/view?usp=sharing).

# Tecnologias utilizadas
- Excel VBA (macros integradas à planilha de controle)
- Google Forms
- Google Sheets + Google Apps Script
- Power BI

# Estrutura do projeto
🔹 **Aba 1 — Solicitação de Orçamento**  
- [`Sub EnviarEmailsSolicitacaoDeOrcamento()`](vba/aba1/EnviarEmailsSolicitacaoDeOrcamento.bas): Responsável por enviar e-mails com links personalizados para preenchimento do Google Forms para cada item com status "Solicitar orçamento".
- [`Sub registrarNovoItem()`](vba/aba1/registrarNovoItem.bas): Registra um novo item de solicitação na planilha, a partir dos campos preenchidos pelo solicitante.  

🔹 **Aba 2 — Pedidos Orçados**  
- [`Sub ImportarRespostasDoGoogleForms()`](vba/aba2/ImportarRespostasDoGoogleForms.bas): Importa automaticamente as respostas do formulário para a planilha, gerando novos registros e atualizando o status dos pedidos.

🔹 **Aba 3 — Pedidos Aprovados**  
- [`Sub TransferirPedidosAprovados()`](vba/aba3/TransferirPedidosAprovados.bas): Transfere os pedidos aprovados da aba anterior para esta aba, com controle por Ticket ID.
- [`Sub MarcarPedidoComoRecebido()`](vba/aba3/MarcarPedidoComoRecebido.bas): Permite que o usuário registre o recebimento de um pedido com base no Ticket ID.

🔹 **Aba 4 — Estoque**  
- [`Sub TransferirParaEstoque()`](vba/aba4/TransferirParaEstoque.bas): Move os itens recebidos para o controle de estoque, garantindo não duplicidade pelo Ticket ID.
- [`Sub RegistrarBaixaEstoque()`](vba/aba4/RegistrarBaixaEstoque.bas): Exibe um formulário para registrar baixas no estoque (retiradas de materiais).
- [`Private Sub btnConfirmar_Click()`](vba/aba4/btnConfirmar_Click.bas): Lógica do botão "Confirmar" no UserForm de baixa, que valida o saldo disponível e registra a saída.

🔹 **Macros Gerais (todas as abas)**  
- [`Sub ApagarLinhasFinais()`](vba/macros-gerais/ApagarLinhasFinais.bas): Remove linhas vazias residuais da tabela.
- [`Sub telacheia()`](vba/macros-gerais/telacheia.bas) / [`Sub telanormal()`](vba/macros-gerais/telanormal.bas): Ajustam o modo de visualização da planilha (tela cheia / normal).

🔹 **Google Apps Script**  
- [`Gerador de ticket ID.gs`](google-apps-script/Gerador_TicketID.gs): Código executado automaticamente após o envio de cada resposta no Google Forms. Gera um código identificador único (Ticket ID) para rastrear o pedido ao longo do fluxo.

# Resultado esperado
Esse sistema proporciona:
- Automação do fluxo de compras;
- Padronização dos registros;
- Maior rastreabilidade dos pedidos;
- Controle em tempo real do estoque;
- Visualização gerencial através do Power BI.
