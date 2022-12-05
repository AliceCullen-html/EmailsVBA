Attribute VB_Name = "multEmails"

sub EnviarEmails ()

set objeto_outlook = CreateObject ("Outlook.Application")


for linha = 2 to range("A1").End(xlDown).Row

    'Criar novo email ' 

    set email = objeto_outlook.createitem(0)

    'Mostrar janela do outlook

    email.Display

    email.to = cells(linha, 1).value
    
    email.Subject = "Relatório de Vendas"

    email.HtmlBody = "Olá" & cells(linha, 2).value & ", " & cells(linha, 3).value & "Um abraço, "

    email.Attachments.Add (ThisWorkbook.Path & "\Relatórios Vendas\Vendas - " & cells(linha, 2).value & ".xlsx")



    email.send



Next



end sub