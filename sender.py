# Send an HTML email with an embedded image and a plain text message for
# email clients that don't want to display the HTML - It prevents phishing mark.

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
import smtplib
import pandas as pd
import progressbar
import getpass

################ Email fixed, password Fixed : #####################
# SenderAddress = "XYZ@gmail.com"
# password = "XXXXXXXXXXX"

################ Email fixed, input password : #####################
#input through cmd...
SenderAddress = "XYZ@gmail.com"
password = getpass.getpass("Enter your secret password : ")

################ Input Email, Input password : #####################
#input through cmd...
# SenderAddress = input("Enter Sender\'s mail address : ")
# password = getpass.getpass("Enter your secret password : ")


############ Read the email directory file with emails and names #########
e = pd.read_excel("email.xlsx")
emails = e['Emails'].values
names = e["Names"].values
print(f"The receiver's mail-ids are : \n\n{emails}")

#Create SMTP session for sending the mail
try :
	smtp = smtplib.SMTP('smtp.gmail.com', 587) #use gmail with port
	smtp.starttls() #enable security
	smtp.login(SenderAddress, password) #login with mail_id and password
except smtplib.SMTPAuthenticationError:
	sys.exit("SMTP User authentication error, Email not sent!")


############## mail body : html ####################

# Short example : 
# We reference the image in the IMG SRC attribute by the ID we give it below
# m = """\
# <b>Some <i>HTML</i> text</b> and an image.<br>
# <img src="cid:image1" width="800px" height="1200px"><br>
# Nifty!
# """

mail_content = """\
<div>Hi {},</div>
<div>&nbsp;</div>
<div>This is a mail with HTML body as MIME text type.&nbsp;<u>This is underlined Text</u></div>
<div>&nbsp;</div>
<div>1. This is point in ordered list.</div>
<div>&nbsp;</div>
<div>2. This is point number 2 with <strong>&nbsp;strong text style.</strong></div>
<div>&nbsp;</div>
<div><img src="cid:image1";disp=emb" width="470" height="711" data-image-whitelisted="" /></div>
<div>&nbsp;</div>
<div>If you find this interesting you can reply to this mail or reach out to the founder directly at&nbsp;<u>98XXXXXX12</u>.</div>
<div>&nbsp;</div>
<div>&nbsp;</div>
<div>Regards,</div>
<div>Atul,</div>
<div><a href="https://python.org/" target="_blank" rel="noopener" data-saferedirecturl="https://www.google.com/url?q=https://python.org/</a></div>
"""

############# Loading IMage embedding ####################
# It assumes the image is in the current directory
fp = open('banner.jpg', 'rb')
msgImage = MIMEImage(fp.read())
fp.close()


# Initializing mail objects, passing mail ids and body for each recepient.

for email, name, i in zip(emails, names, progressbar.progressbar(range(len(emails)))):
	# Create the root message and fill in the from, to, and subject headers
	msgRoot = MIMEMultipart('related')
	msgRoot['Subject'] = 'This is the subject of your mail.'
	msgRoot['From'] = SenderAddress
	msgRoot['To'] = email
	msgRoot.preamble = 'This is a multi-part message in MIME format.'

	# Encapsulate the plain and HTML versions of the message body in an
	# 'alternative' part, so message agents can decide which they want to display.
	msgAlternative = MIMEMultipart('alternative')
	msgRoot.attach(msgAlternative)

	msgText = MIMEText('This is the alternative plain text message.')
	msgAlternative.attach(msgText)


	msgText = MIMEText(mail_content.format(name), 'html') #this prohibts from phishing mark.
	msgAlternative.attach(msgText)


	# Define the image's ID as referenced above
	msgImage.add_header('Content-ID', '<image1>')
	msgRoot.attach(msgImage)

	# send your mail
	smtp.sendmail(SenderAddress, email, msgRoot.as_string())

# Closing the smtp server session...
smtp.quit()
print("\n\nAll mails have been sent successfully...!")