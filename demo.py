import smtplib, ssl
import pandas as pd
from os.path import basename
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

message = MIMEMultipart()

# start the email login part
smtp = smtplib.SMTP("smtp.outlook.com")
smtp.starttls()
smtp.login("aashinde20@gmail.com", "Satara@123")
print("lohin success")
# email login part end

# reading excel file
email_lists = pd.read_excel("./new.xlsx")

names = email_lists["NAME"]
emails = email_lists["EMAIL"]
status = email_lists["STATUS"]


for i in range(len(emails)):
    name = names[i]
    email = emails[i]
    stts = status[i]

    # MEssage format we have to send to the particular user
    text = (
        "Your report for the payment of last quarter \nName: "
        + name
        + " \nPayment Status: "
        + stts
    )
    # print(text)
    # html = """\
    # <html>
    # <body>
    #     <table border="1">
    #         <tr>
    #             <th>NAME</th>
    #             <th>EMAIL</th>
    #             <th>STATUS</th>
    #         </tr>
    #         <tr>
    #             <td>{% name %}</td>
    #             <td>$email</td>
    #             <td>{{ status }}</td>
    #         </tr>
    #     </table>

    # </body>
    # </html>
    # """
    part2 = MIMEText(text, "plain")
    # part2 = MIMEText(html, "html")

    # message.attach(part2)

    if stts == "Overdue":
        message = MIMEMultipart("alternative")
        message["Subject"] = "General Notice"
        message["From"] = "aashinde20@gmail.com"
        message["To"] = email
        message["CC"] = "sanketj019@gmail.com"
        message.attach(part2)
        smtp.sendmail("aashinde20@gmail.com", [email], message.as_string())
        print(i)
    else:
        print("No overdues")

print("Email has been sent to ")
