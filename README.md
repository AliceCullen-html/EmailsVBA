# Integração Outlook com VBA 📧

# Criando objeto 

Para criar a caixa de e-mail , onde vamos preencher todas as informações necessárias para enviar os e-mails, é necessário criaro objeto Email, por meio da
propriedade <b> CreateItem</b>.

![image](https://user-images.githubusercontent.com/77951123/205712328-b44a272a-d593-4163-b382-f4d8c5af52c7.png)

Além disso, também será necessárioo comando <b>.Display</b> para que seja possível visualizar a caixa de e-mail.

<hr>

# Propriedades do Corpo do E-mail

![image](https://user-images.githubusercontent.com/77951123/205712663-5260f28a-3466-4f0c-9116-098f6f1c8a34.png)

<hr>

# Enviando e-mail simples

![image](https://user-images.githubusercontent.com/77951123/205712805-a5eec292-bbb8-41b7-bd33-ee8660a4b07f.png)

# Enviando vários e-mail a partir de lista no Excel

Como queremos enviar o e-mail repetidas vezes, então precisamos
criar uma estrutura de repetição, percorrendo a linha 2 até a última
linha preenchida da tabela:

![image](https://user-images.githubusercontent.com/77951123/205713118-585c7cae-2500-4277-97c8-d105c158da57.png)

<hr>

# Colocando anexos


![image](https://user-images.githubusercontent.com/77951123/205713209-33c953b1-a38c-4981-95ec-800330a619ca.png)

Basta colocar o caminho do arquivo no **Email.Attachmentes.Add**, 
podendo ainda concatenar dados do excel ao caminho do arquivo como no exemplo acima.





<img align="center" src="http://ForTheBadge.com/images/badges/built-with-love.svg" />
