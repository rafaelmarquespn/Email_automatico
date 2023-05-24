# Como utilizar

1. Copiar a pasta app para sua ide
2. Logar no aplicativo do outlook no seu pc localmente
3. Baixar socket do windows para comunicar com outlook
   ```
   pip install pywin32
   ```

Depois desse passo a passo, vá para pasta archives:

1. Coloque os emails que você pretende atingir no usuarios_destino.txt
2. 1 por linha
3. Depois faça o email com html
4. E o salve em corpo_email.html


Vá para o arquivo email_dump.py:

1. Coloque o assunto que você quiser na variável Subject
2. Se existir algum anexo coloque na variável anexo
3. Ready to run!


Execute o arquivo e espere o log de envio com sucesso
