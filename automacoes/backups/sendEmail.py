#imports send email 
import smtplib
import email.message
import datetime

#send class function 
class EmailConf:
    def __init__(self):
        self.email1 = 'jvcbcarvalho@gmail.com'
        self.senha = 'lhkl neir pmrt tkzx'

    def enviaemail(self, destinatario, dest_copia, assunto, corpo):
        hora_atual = datetime.datetime.now().time()
        manha_inicio = datetime.time(6, 0, 0)  # 6:00 - morning 
        tarde_inicio = datetime.time(16, 0, 0)  # 16:00 - afternoon  
        noite_inicio = datetime.time(18, 0, 0)  # 23:00 - afternoon 

        if hora_atual < manha_inicio:
            saudacao = "Boa noite"
        elif hora_atual < tarde_inicio:
            saudacao = "Bom dia"
        elif hora_atual < noite_inicio:
            saudacao = "Boa tarde"
        else:
            saudacao = "Boa noite"
        try:
            corpo_email = f"""
            <p>{saudacao}!</p>
            <p></p>
            <p></p>
            <p>{corpo}</p>
            <p>Lista de execuções e horários: </p>
            <p>-> 16:30 - ctccVuca </p>
            <p>-> 17:00 - spqVuca</p>
            <p>-> 23:30 - ctccVuca </p>
            <p>-> 00:00 - spqVuca</p>
            <p>-> 02:00 - spqVeiChicoZigpay e spqManeZigpay</p>
            <p>Qualquer apontamento indevido, favor sinalizar.</p>
            <p></p>
            <p>--<br>
            João Barros<br>
            Full Stack Developer<br>
            <br>
            Contato - 61 9 9252 5843<br>
            Email: jvcbcarvalho@gmail.com</p>
            <img src="https://i.imgur.com/KLUHTR8.png heigt="200" width="300" alt="João Barros">
            """

            msg = email.message.Message()
            msg['Subject'] = assunto
            msg['From'] = self.email1
            msg['To'] = destinatario
            msg['Cc'] = dest_copia
            msg.add_header('Content-Type', 'text/html')
            msg.set_payload(corpo_email)

            with smtplib.SMTP('smtp.gmail.com', 587) as s:
                s.starttls()  # Habilita o TLS
                s.login(self.email1, self.senha)

                destinatarios = destinatario.split(', ') + dest_copia.split(', ')

                s.sendmail(msg['From'], destinatarios, msg.as_string().encode('utf-8'))
                print('E-mail enviado com sucesso!')

        except smtplib.SMTPAuthenticationError:
            print('Erro de autenticação. Verifique o endereço de e-mail e senha.')
        except smtplib.SMTPException as e:
            print(f'Erro ao enviar o e-mail: {e}')

# Example usage
email_sender = EmailConf()
email_sender.enviaemail('jvcbcarvalho@gmail.com', 'felipe@orionbi.com.br', 'Alerta de execução de automação', 'Confirmação de execução.')
