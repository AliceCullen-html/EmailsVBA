Attribute VB_Name = "Emailvb"

sub EnviarEmail ()

'Criar objeto dentro de var'

set objeto_outlook = CreateObject ("Outlook.Application")

'Criar novo email ' 

set email = objeto_outlook.createitem(0)

'Mostrar janela do outlook

email.Display

email.to = "krattosgamer@hotmail.com"
email.cc = "marcusvcunha0800@hotmail.com"

email.Subject = "Testando"

email.HtmlBody = "Testando e-mail via vba"


email.send







end sub 
