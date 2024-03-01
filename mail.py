import openpyxl
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime, timedelta

months = {
    1: "Ocak",
    2: "Subat",
    3: "Mart",
    4: "Nisan",
    5: "Mayis",
    6: "Haziran",
    7: "Temmuz",
    8: "Ağustos",
    9: "Eylül",
    10: "Ekim",
    11: "Kasim",
    12: "Aralik"
}#Datetime ile now.strftime("%B") bu kodu kullandığınız zaman aylar ingilizce olarak döner, 
#Türkçe gelmesini istediğim için bu şekilde bir çeviri disi oluşturdum

#Tarih değişkenlerini aldığımız kısım
now = datetime.now()
sheetNumber = now.day-1
monthIndex = now.month
month=months[monthIndex]


# Excel dosyasını aç
filename = month+"Ornek-Gunluk-Rapor.xlsx"
workbook = openpyxl.load_workbook(filename)
sheets = workbook.sheetnames

# İstenen tabloya eriş
sheet = workbook[sheets[sheetNumber]]

# E-posta göndermek için gerekli bilgiler
my_email = 'your_email@example.com' # Kendi e-posta adresinizi buraya yazın
my_password = 'your_password' # E-posta şifrenizi buraya yazın 

#Mail ayarlarınız üzerinden 3.parti uygulamalara izin verin,aksi takdirde authentication hatası alırsınız.

smtp_server = 'smtp.example.com' # SMTP sunucu adresini buraya yazın
smtp_port = 587
receiver_email='receiver_email@example.com' # Alıcı e-posta adresini buraya yazın

# Dosya eklemek için
attachment = open(filename, "rb")

cellValue = sheet['B2'].value #Rapor yazarken B2 sütunu kesin dolu olacağı için buradan kontrol sağlıyorum

if cellValue: #B2 sütununda herhangi bir değer varsa

    #Eposta oluşturduğumuz kısım
    msg = MIMEMultipart()
    msg['From'] = my_email
    msg['To'] = receiver_email
    msg['Subject'] = 'Günlük Rapor'
    body = "Merhaba X bey/hanım,\nİstemiş olduğunuz rapor ekte mevcuttur.\nİyi akşamlar dilerim. "
    msg.attach(MIMEText(body, 'plain'))
    part = MIMEBase('application', 'octet-stream')
    part.set_payload((attachment).read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
    msg.attach(part)

    #Epostayı gönderdiğimiz alan
    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(my_email, my_password)
        text = msg.as_string()
        server.sendmail(my_email, receiver_email, text)
        server.quit()
        print("E-posta başarıyla gönderildi.")
    except Exception as e:
        print("E-posta gönderilirken bir hata oluştu:", e)
else:
    msg = MIMEMultipart()
    msg['From'] = my_email
    msg['To'] = receiver_email
    msg['Subject'] = 'Günlük Rapor'
    body = "Merhaba X bey/hanım,\nBugün SLA dışı herhangi bir istek gelmemiştir. Tarafınıza bildirme amaçlı göndermiş olduğum maildir.\nİyi akşamlar dilerim. "
    msg.attach(MIMEText(body, 'plain'))
    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(my_email, my_password)
        text = msg.as_string()
        server.sendmail(my_email, receiver_email, text)
        server.quit()
        print("E-posta başarıyla gönderildi.")
    except Exception as e:
        print("E-posta gönderilirken bir hata oluştu:", e)

attachment.close()
