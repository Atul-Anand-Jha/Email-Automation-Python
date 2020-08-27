from email.message import EmailMessage
from email.utils import make_msgid
import mimetypes
import smtplib

msg = EmailMessage()

SenderAddress = "XYZ@gmail.com"
password = "ndXX@XX3$#XXX"

# Define these once; use them twice!
strFrom = SenderAddress
strTo = 'abc.123@gmail.com'

# generic email headers
msg['Subject'] = 'Hello there'
msg['From'] = SenderAddress

# set the plain text body
msg.set_content('This is a plain text body.')

# now create a Content-ID for the image
image_cid = make_msgid(domain='testDomain.com')
# if `domain` argument isn't provided, it will 
# use your computer's name

# set an alternative html body
msg.add_alternative("""\
<html>
    <body>
        <p>This is an HTML body.<br>
           It also has an image.
        </p>
        <img src="cid:{image_cid}">
    </body>
</html>
""".format(image_cid=image_cid[1:-1]), subtype='html')
# image_cid looks like <long.random.number@xyz.com>
# to use it as the img src, we don't need `<` or `>`
# so we use [1:-1] to strip them off


# now open the image and attach it to the email
with open('banner.jpg', 'rb') as img:

    # know the Content-Type of the image
    maintype, subtype = mimetypes.guess_type(img.name)[0].split('/')

    # attach it
    msg.get_payload()[1].add_related(img.read(), 
                                         maintype=maintype, 
                                         subtype=subtype, 
                                         cid=image_cid)


# the message is ready now
# you can write it to a file
# or send it using smtplib

smtp = smtplib.SMTP('smtp.gmail.com', 587, timeout=15)
smtp.starttls()
smtp.login(SenderAddress, password)
smtp.sendmail(strFrom, strTo, msg.as_string())
smtp.quit()
print("Main Sent! ")