import smtplib
from email.mime.text import MIMEText

def main_send_email (addressee,subject,message,sender, password):
    server = smtplib.SMTP("smtp.gmail.com", 587) # Обращение к серверу отправки почты попорту 587
    server.starttls()  # Запуск обменна даных
    try:
        server.login(sender, password)  # Регистрация на сервере
        msg = MIMEText(message, 'html')  # Конвертация в html формат
        msg["From"] = sender #От кого
        msg["To"] = addressee #Ккому
        msg["Subject"] = subject  # Темма сообщения
        server.sendmail(sender, addressee, msg.as_string())  # Отправка письма от кого кому и сообщение
        return f'Ок'
    except:
        return f' При отправке письма {addressee} произошла ошибка. Письмо не отправлено.'




if __name__ == '__main__':
    print(main_send_email('timofei.kropachev@gmail.com','Тема письма', 'web\index_1.html','timofei.kropachev@gmail.com','oqtspbctwzowkbfm'))
  