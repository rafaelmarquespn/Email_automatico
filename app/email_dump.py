import win32com.client as win32


class EmailDumps():
    def __init__(self, subject, anexo = None):
        self.subject = subject
        self.anexo = anexo
        self.body = self.load_data()
        self.destinatarios = self.load_destinatarios()

    def load_data(self):
        with open("C:\\dev\\projetos\\EmailAutomatico\\app\\archives\\corpo_email.html", "r", encoding="utf-8") as f:
            data = f.readlines()
            body = ""
            for line in data:
               body += line
        return body
    
    def load_destinatarios(self):
        with open("C:\\dev\\projetos\\EmailAutomatico\\app\\archives\\usuarios_destino.txt", "r", encoding="utf-8") as f:
            data = f.readlines()
            destinatarios = ""
            for line in data:
                destinatarios += line.strip() + ';'
        return destinatarios
    
    def send_email(self):
        outlook = win32.Dispatch("Outlook.Application")
        email = outlook.CreateItem(0)
        email.To = self.destinatarios
        email.Subject = self.subject
        email.HTMLBody = self.body
        if self.anexo:
            email.Attachments.Add(self.anexo)

        email.Send()
        print("E-mail enviado com sucesso!")


#colocar assunto do e-mail
Subject = "E-mail automático python"

#colocar o caminho do arquivo
anexo = None
#cria objeto
if "__name__ " == "__main__":
    email = EmailDumps(Subject, anexo)
    email.send_email()



