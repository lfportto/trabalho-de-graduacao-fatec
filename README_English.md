# Purchase Request, Receiving, and Inventory Management System
📄 [Versão em português](README.md)
## About
This repository contains the code developed in `VBA` (Visual Basic for Applications) and `Google Apps Script` as part of the Graduation Project presented to the Faculdade de Tecnologia de São José dos Campos – Prof. Jessen Vidal (FATEC SJC), during the 5th and 6th semesters of the Industrial Production Management Technology course, by students Luis Felipe Porto and Rodrigo da Silva Oliveira.

The project aims to automate and standardize the process of purchase requests, quotations, approvals, receiving, and inventory control of materials for a small-sized company, through a spreadsheet integrated with forms and a management dashboard for monitoring and interaction in Power BI.

The complete workflow is described in the Graduation Project [available here](https://drive.google.com/file/d/1il2iBtzbF8Q_8AwS4Z1RSmsinUmUGz2x/view?usp=sharing).

## Technologies Used
- Excel VBA (macros integrated into the control spreadsheet)
- Google Forms
- Google Sheets + Google Apps Script
- Power BI

## Project Structure
🔹 **Sheet 1 — Budget Request**  
- [`Sub EnviarEmailsSolicitacaoDeOrcamento()`](vba/aba1/EnviarEmailsSolicitacaoDeOrcamento.bas): Responsible for sending emails with personalized links to fill out the Google Forms for each item with status "Request budget".  
- [`Sub registrarNovoItem()`](vba/aba1/registrarNovoItem.bas): Registers a new request item in the spreadsheet, based on the fields filled by the requester.  

🔹 **Sheet 2 — Quoted Orders**  
- [`Sub ImportarRespostasDoGoogleForms()`](vba/aba2/ImportarRespostasDoGoogleForms.bas): Automatically imports form responses into the spreadsheet, generating new records and updating the status of orders.

🔹 **Sheet 3 — Approved Orders**  
- [`Sub TransferirPedidosAprovados()`](vba/aba3/TransferirPedidosAprovados.bas): Transfers approved orders from the previous sheet to this one, controlling by Ticket ID.  
- [`Sub MarcarPedidoComoRecebido()`](vba/aba3/MarcarPedidoComoRecebido.bas): Allows the user to register the receipt of an order based on the Ticket ID.

🔹 **Sheet 4 — Inventory**  
- [`Sub TransferirParaEstoque()`](vba/aba4/TransferirParaEstoque.bas): Moves received items to inventory control, ensuring no duplication by Ticket ID.  
- [`Sub RegistrarBaixaEstoque()`](vba/aba4/RegistrarBaixaEstoque.bas): Displays a form to register inventory withdrawals (material removals).  
- [`Private Sub btnConfirmar_Click()`](vba/aba4/btnConfirmar_Click.bas): Logic for the "Confirm" button in the withdrawal UserForm, which validates available balance and records the output.

🔹 **General Macros (all sheets)**  
- [`Sub ApagarLinhasFinais()`](vba/macros-gerais/ApagarLinhasFinais.bas): Removes residual empty rows from the table.  
- [`Sub telacheia()`](vba/macros-gerais/telacheia.bas) / [`Sub telanormal()`](vba/macros-gerais/telanormal.bas): Adjust the spreadsheet display mode (full screen / normal).

🔹 **Google Apps Script**  
- [`Gerador de ticket ID.gs`](google-apps-script/Gerador_TicketID.gs): Code automatically executed after each Google Forms submission. Generates a unique identifier code (Ticket ID) to track the order throughout the workflow.

## Expected Outcome
This system provides:
- Automation of the purchase flow;  
- Standardization of records;  
- Greater traceability of orders;  
- Real-time inventory control;  
- Management visualization through Power BI.
