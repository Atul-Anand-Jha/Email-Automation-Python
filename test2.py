import pandas as pd
import smtplib
import imghdr
from email.message import EmailMessage

SenderAddress = "XYZ@gmail.com"
password = "ndXX@XX3$#XXX"

e = pd.read_excel("email.xlsx")
emails = e['Emails'].values
names = e["Names"].values
file = "banner.jpg"
msg = EmailMessage()
msg['Subject'] = "Hello world - dynamic"
msg['From'] = SenderAddress
print(f"The receiver's mail ids are : \n\n{emails}")

with smtplib.SMTP("smtp.gmail.com", 587, timeout=15) as server:
	server.starttls()
	server.login(SenderAddress, password)
	# msg = f"Hello {this is an email form python"
	# subject = "Hello world"
	# body = "Subject: {}\n\n{}".format(subject,msg)
	with open(file, 'rb') as f:
		file_data = f.read()
		file_type = imghdr.what(f.name)
		file_name = f.name

	for email,name in zip(emails,names):
		
		msg['To'] = email
		
		body = f"Hello {name};\n\n\nThis is an email from python"
		# msg = "Subject: {}\n\n{}".format(subject,body)
		msg.set_content(body)
		msg.add_attachment(file_data, maintype='image', subtype=file_type, filename=file_name)
		server.send_message(msg)
		# server.sendmail(SenderAddress, email, msg)
	server.quit()
