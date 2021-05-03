# Criando_e_Enviando_Relatorio_de_Vendas

========= Enviando Relatório de Vendas de lojas por e-mail (usando outloook) =========


Esse código é uma adaptação minha do aprendizado do Minicurso Python da hashtag programação.
link do curso: https://pages.hashtagtreinamentos.com/minicurso-python-automacao?blog=1n4033rer&video=3dep762tr



Requisitos: 
• instalar as biliotecas do python (pandas, win32com e pyautogui) caso não tenha.
    . para isso use o cmd e rode esses comandos
        pip install pandas
        pip install win32com
        pip install pyautogui
• Ter o outlook instalado e configurado no email do remetente.



Código pronto para enviar qualquer email pelo outlook:

outlook = win32.Dispatch('outlook.application') # seconectando como outlook do pc
mail = outlook.CreateItem(0) # criando email
mail.To = 'dudahmovies@gmail.com' # pra que enviar
mail.Subject = 'Assunto do email' # assunto do email
mail.HTMLBody = 'Messagem' # No corpo do email pode ser usado código html
mail.Send() # enviar email
