import gspread, smtplib, datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Гайд по получению этого файла тут: "https://www.geeksforgeeks.org/how-to-automate-google-sheets-with-python/"
gc = gspread.service_account(filename='your_files_name.json')
sh = gc.open("Sheets_name")
worksheet = sh.sheet1


# Ищет последнюю свободную строку в таблице
def next_available_row(worksheet):
    str_list = list(filter(None, worksheet.col_values(1)))
    return len(str_list) + 1


# Отправка сообщений на почту
def send_mail(for_mail, to):
    # Позволяет отправлять русские сообщения
    msg = MIMEMultipart()
    # Пароль получается через настройки безопасности в почте
    # "https://account.mail.ru/user/2-step-auth/passwords?back_url=https%3A%2F%2Fid.mail.ru%2Fsecurity"
    password = 'your_password'
    # Почта через которую отправляются письма
    msg['From'] = "your_mail"
    msg['To'] = to
    # Заголовок
    msg['Subject'] = "subject"
    msg.attach(MIMEText(for_mail, 'plain'))
    server = smtplib.SMTP('smtp.mail.ru: 587')
    server.starttls()
    server.login(msg['From'], password)
    server.sendmail(msg['From'], msg['To'], msg.as_string())
    server.quit()
    # Запись в log кому было отправлено письмо
    logging = open('log.txt', 'a')
    logging.write(str(datetime.datetime.now()) + ' Сообщение было успешно отправлено: %s' % (msg['To']) + '\n')
    logging.close()
    print("successfully sent email to %s:" % (msg['To']))


# Цикл берет нужные информации из таблицы
for i in range(next_available_row(worksheet) - 2):
    i += 1
    name = worksheet.cell(i + 1, 1).value
    mail = worksheet.cell(i + 1, 2).value
    flag = worksheet.cell(i + 1, 3).value
    point = worksheet.cell(i + 1, 4).value
    place = worksheet.cell(i + 1, 5).value
    if flag == 'Да':
        try:
            # Выбираем из какого файла будет взят текст
            file = open('Accepted.txt', 'r')
            # Подставляем в фигурные скобки
            for_mail = file.read().format(name, place, point)
            send_mail(for_mail=for_mail, to=mail)
        except Exception as _ex:
            # Запись в log ошибки
            logging = open('log.txt', 'a')
            logging.write(str(datetime.datetime.now()) + f"{_ex}\nCheck your login or password" + '\n')
            logging.close()
            print(f"{_ex}\nCheck your login or password")
    else:
        try:
            # Выбираем из какого файла будет взят текст
            file = open('Declined.txt', 'r')
            # Подставляем в фигурные скобки
            for_mail = file.read().format(name=name, point=point)
            send_mail(for_mail=for_mail, to=mail)
        except Exception as _ex:
            # Запись в log ошибки
            logging = open('log.txt', 'a')
            logging.write(str(datetime.datetime.now()) + f"{_ex}\nCheck your login or password" + '\n')
            logging.close()
            print(f"{_ex}\nCheck your login or password")
