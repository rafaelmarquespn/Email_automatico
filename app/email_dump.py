import win32com.client as win32


class EmailDumps():
    def __init__(self, subject, anexo = None):
        self.subject = subject
        self.anexo = anexo
        self.body = self.load_data()
        self.destinatarios = self.load_destinatarios()
        self.usuario_destino = None

    def load_data(self):
        with open("C:\\dev\\projetos\\EmailAutomatico\\app\\archives\\corpo_email.html", "r", encoding="utf-8") as f:
            data = f.readlines()
            body = ""
            for line in data:
               line = line.strip()
               body += line
        return body
    
    def load_destinatarios(self):
        with open("C:\\dev\\projetos\\EmailAutomatico\\app\\archives\\usuarios_destino.txt", "r", encoding="utf-8") as f:
            data = f.readlines()
            destinatarios_lista = []
            for line in data:
                destinatarios = line.strip()
                destinatarios_lista.append((destinatarios.split("; ")[0], destinatarios.split("; ")[1]))

        return destinatarios_lista
    
    def send_email(self):
        body = self.body
        outlook = win32.Dispatch("Outlook.Application")
        email = outlook.CreateItem(0)
        email.Subject = self.subject
        for destinatario in self.destinatarios:
            email.To = destinatario[0]
            self.usuario_destino = destinatario[1]
            body = body.replace("usuario_destino", self.usuario_destino)

            email.HTMLBody = body
            if self.anexo:
                email.Attachments.Add(self.anexo)

            email.Send()
            print(f"E-mail enviado com sucesso para @{destinatario[1]}!")


#colocar assunto do e-mail
Subject = "E-mail automático python"
#colocar o caminho do arquivo
anexo = None
#cria objeto



if __name__  == "__main__":
    email = EmailDumps(Subject, anexo)
    email.send_email()



