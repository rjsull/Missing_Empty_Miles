# Python code to illustrate Sending mail with attachments 
# from your Gmail account  
  
# libraries to be imported 
import os
import smtplib 
from email.mime.multipart import MIMEMultipart 
from email.mime.text import MIMEText
#from email.mime.image import MIMEImage
from email.mime.base import MIMEBase 
from email import encoders 

# Put your email for fromaddr and add recipients for toaddr and bcc
fromaddr = 'x'
toaddr = ['x']
bcc = ['x', 'x', 'x']

# instance of MIMEMultipart 
msg = MIMEMultipart() 
  
# storing the senders email address   
msg['From'] = fromaddr 
  
# storing the receivers email address  
msg['To'] = ", ".join(toaddr) 
  
# storing the subject  
msg['Subject'] = "Missing Empty Miles"
  
# string to store the body of the mail, includes a hyperlink 
body = """<pre> <font face="calibri" color="black" size="3">
All,

Attached is the missing empty miles summary. 

Thanks,
--

</font></pre>"""

# attach the body with the msg instance 
msg.attach(MIMEText(body, 'html')) 

#set working directory 
os.chdir('C:\\Users\\sullivanry\\Documents')

# attach the image
#fp = open('cust_miles_wrap.jpg', 'rb')
#msgImage = MIMEImage(fp.read())
#fp.close()

# Define the image's ID as referenced above
#msgImage.add_header('Content-ID', '<image1>')
#msg.attach(msgImage)

# open the file to be sent, has to be in working directory, change file to invalid if you don't need attachments 
filename = "MissingEmptyMiles.xlsx"
attachment = open(filename, "rb") 

# instance of MIMEBase and named as p 
p = MIMEBase('application', 'octet-stream') 

# To change the payload into encoded form 
p.set_payload((attachment).read()) 

# encode into base64 
encoders.encode_base64(p) 
p.add_header('Content-Disposition', "attachment; filename= %s" % filename) 
# attach the instance 'p' to instance 'msg' 
msg.attach(p) 
# creates SMTP session 
s = smtplib.SMTP('smtp.gmail.com', 587) 
  
# start TLS for security 
s.starttls() 

# Authentication, Enter Your Password Below where the x is
s.login(fromaddr, "x") 
  
# Converts the Multipart msg into a string 
text = msg.as_string() 
  
# sending the mail 
s.sendmail(fromaddr, toaddr, text) 
  
# terminating the session 
s.quit() 
