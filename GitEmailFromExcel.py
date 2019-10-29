import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase 
from email import encoders 
import pandas as pd


def sendEmail(sender_email,receiver_email,cc,Company_Name,text,html,password):
    with smtplib.SMTP('smtp.gmail.com',587) as server:
        server.ehlo()
        server.starttls()
        server.ehlo()

        #set up multipart message
        message = MIMEMultipart("alternative")
        message["Subject"] = f"Hack the Burgh VI - {Company_Name}"
        message["From"] = sender_email
        message["CC"] = cc
        message["To"] = receiver_email

        # Turn text and html into plain/html MIMEText objects
        part1 = MIMEText(text, "plain")
        part2 = MIMEText(html, "html")

        filename = "htb2020_prospectus.pdf"
        attachment = open(filename, "rb") 
  
        # instance of MIMEBase and named as p 
        p = MIMEBase('application', 'octet-stream') 
        
        # To change the payload into encoded form 
        p.set_payload((attachment).read()) 
        
        # encode into base64 
        encoders.encode_base64(p) 
        
        p.add_header('Content-Disposition', "attachment; filename= %s" % filename) 
        
        # attach the instance 'p' to instance 'msg' 
        message.attach(p) 

        # Add HTML/plain-text parts to MIMEMultipart message
        # The email client will try to render the last part first
        message.attach(part1)
        message.attach(part2)

        server.login(sender_email, password)
        server.sendmail(sender_email,receiver_email,message.as_string())
        server.quit()

    print("Message sent!!!")

def makeTextEmail(Company_Name,companys_product_or_API,your_FULL_name,your_title,personal_story):
    if personal_story == "None":
        personal_story = ""
    else:
        personal_story = "\n"+personal_story+"\n"
    text = f"""\
    Dear {Company_Name},

    We’re reaching out because we’re beginning to offer limited sponsorship opportunities for companies we think would be a good fit for our audience and we would love to have {Company_Name} as one of our main sponsors.

    Our goal is to give our talented developers a chance to interact with the best companies around Edinburgh and beyond. We would love to partner with you in order to get {companys_product_or_API} out in front of the top developers in the country and give them a chance to work on challenges set by {Company_Name} itself, so that they could explore and learn about what you do in a more hands-on manner. 
    {personal_story}
    We have attached our prospectus with more information about our event, sponsorship packages and benefits. Please do not hesitate to reply to this email if you have any further questions.

    Thank you for your time and we are looking forward to working with you.

    All the best,

    {your_FULL_name}
    Hack the Burgh VI {your_title}
    """
    return text

def makeHTMLEmail(Company_Name,companys_product_or_API,your_FULL_name,your_title,website_link,personal_story = None):
    if personal_story == "None":
        personal_story = ""
    else:
        personal_story = "<p>"+personal_story+"</p>"
    html = f"""\
        <html>
        <body >
            <h1 style="color: rgba(2, 0, 146, 0.801)">Dear {Company_Name}</h1>

            <p >We’re reaching out because we’re beginning to offer limited sponsorship opportunities for companies we think would be a good fit for our audience and we would love to have <b>{Company_Name}</b> as one of our main sponsors.</p>

            <p>Our team at Hack the Burgh is organising <b>Scotland’s largest 24-hour hackathon, with over 250 skilled tech developers</b> from over 30 universities. We are happy to announce that we are coming back with the sixth edition on the <b>29th of February 2020 to the 1st of March 2020</b>.</p>


            <p>Our goal is to give our talented developers a chance to interact with the best companies around Edinburgh and beyond. We would love to partner with you in order to get <b>{companys_product_or_API}</b> out in front of the top developers in the country and give them a chance to work on challenges set by {Company_Name} itself, so that they could explore and learn about what you do in a more hands-on manner. </p>

            {personal_story}

            <p>We have attached our prospectus with more information about our event, sponsorship packages and benefits at <a href="{website_link}">hacktheburgh.com</a>. Please do not hesitate to reply to this email if you have any further questions.</p>

            <p>Thank you for your time and we are looking forward to working with you.</p>
            <p>All the best,</p>
            <div style="font-weight: bold;">{your_FULL_name}</div>
            <div style="font-weight: bold;">Hack the Burgh VI {your_title}</div>
            
        </body>
        </html>
        """
    return html

def emailFromExcelData(filename,in_password):
    sponsorsheet = pd.read_excel(filename,sheet_name = "sponsoremails", header = 0)
    rows,cols = sponsorsheet.shape
    for r in range(rows):
        #fill data from spreadsheet
        sender_email = sponsorsheet.iloc[r,0]
        receiver_email = sponsorsheet.iloc[r,1]
        cc = sponsorsheet.iloc[r,2]
        Company_Name = sponsorsheet.iloc[r,3]
        companys_product_or_API = sponsorsheet.iloc[r,4] #...in order to get ______ out in front of...
        personal_story = sponsorsheet.iloc[r,5]
        your_FULL_name = sponsorsheet.iloc[r,6]
        your_title = sponsorsheet.iloc[r,7] #one of Head of Finance/Financial Officer
        website_link = sponsorsheet.iloc[r,8]

    text = makeTextEmail(Company_Name,companys_product_or_API,your_FULL_name,your_title,personal_story = personal_story)
    print(text)
    html = makeHTMLEmail(Company_Name,companys_product_or_API,your_FULL_name,your_title,website_link,personal_story = personal_story)

    sendEmail(sender_email,receiver_email,cc,Company_Name,text,html,in_password)

    

#PANDAS reading from exccel file

password = input("Enter password to sender account: ")
emailFromExcelData("sponsoremails.xlsx",password)
