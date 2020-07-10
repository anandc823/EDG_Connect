import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

#dummy data
myData = [
    {
        'topic':"Making an awesome Hack Day Project",
        'members':["ageannop@mathworks.com", "aliu@mathworks.com"]
    },
    {
        'topic':"Being great Team Players during Hack Day",
        'members':["ageannop@mathworks.com", "achitale@mathworks.com", "mqureshi@mathworks.com", "skaza@mathworks.com"]
    }
]

#set up email sending constants
address = 'smtp.office365.com'
port = 587
sender = 'ageannop@mathworks.com'
password = #secret!

#send the emails
def sendEmails(serverAddress, serverPort, sendingAddress, sendingPassword, data):

    server = smtplib.SMTP(serverAddress, serverPort)
    server.ehlo()
    server.starttls()
    server.login(sendingAddress, sendingPassword)

    for meeting in data:
        print('parsing this meeting: ' + str(meeting))

        msg = MIMEMultipart()
        msg['From'] = sendingAddress
        msg['To'] = ', '.join(meeting['members']) #converts list of emails into one string, separated by commas
        msg['Subject'] = 'Your EDGConnect Results'

        body = 'The recipients of this email all have an interest in ' +  meeting['topic'] + '.\nYou all should form a group and hang out sometime!'

        msg.attach(MIMEText(body, 'plain'))

        #Still need to generate and attach ics (calendar meeting file) . . .

        text = msg.as_string()
        print(text)
        server.sendmail(sendingAddress, meeting['members'], text)
    #endfor
    server.quit()

sendEmails(address, port, sender, password, myData)

