import pandas as pd
import smtplib

# %%
'''
Change these to your credentials and name
'''
your_name = "Something"
your_email = "Something@gmail.com"
your_password = "Password"

# If you are using something other than gmail
# then change the 'smtp.gmail.com' and 465 in the line below
server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
server.ehlo()
server.login(your_email, your_password)

# %%
# Read the file
email_list = pd.read_excel(r'C:\Users\Gebruiker\Downloads\VariabelBedrag.xlsx')
# %%

# Get all the Names, Email Addreses, Subjects and Messages
all_names = email_list['NAAM']
all_emails = email_list['MAIL']
all_teams = email_list['TEAM']
all_amounts = email_list['BEDRAG']
all_actual = email_list['TEAMACTUAL']
all_expected = email_list['TEAMEXPECTED']
all_topay = email_list['TEBETALEN']
all_expectedtopay = email_list['VERWACHTTEBETALEN']


# %%
# Loop through the emails
for idx in range(len(all_emails)):

    # Get each records name, email, subject and message
    name = all_names[idx]
    email = all_emails[idx]
    subject = "24u Parade Orangade"
    if all_expected[idx] > all_actual[idx]:
        message = "Beste " + all_names[idx] +",\n\n" \
        "Bedankt dat je een van onze teams wilt sponsoren voor onze Parade Orangade. Nu de leiding vele stappen achter " \
        "de rug heeft hopen we ook dat je jouw belofte na komt en de beloofde centjes op onze rekening " \
        "BE95 7765 9995 0858 stort. Plaats in de mededeling je naam en welk(e) team(s) je sponsort. " \
        "Jij sponsorde " + all_teams[idx] + ". Dit team heeft {:.2f} km gewandeld. Jij steunde dit team " \
        "€ {:.2f}  per kilometer. " \
        "Dat maakt dus € {:.2f}.  Eenmaal we dit ontvangen hebben maak je " \
        "ook effectief kans op een van onze mooie prijzen!  Op 15 mei zullen we de prijzen verdelen. " + "\n" + "\n" \
        "Alvast bedankt, \n\nDe leiding"
        message = message.format(all_actual[idx], all_amounts[idx], all_topay[idx])

    else:
        message = "Beste " + all_names[idx] +",\n\n" \
        "Bedankt dat je een van onze teams wilt sponsoren voor onze Parade Orangade. Nu de leiding vele stappen achter " \
        "de rug heeft hopen we ook dat je jouw belofte na komt en de beloofde centjes op onze rekening " \
        "BE95 7765 9995 0858 stort. Plaats in de mededeling je naam en welk(e) team(s) je sponsort. " \
        "Jij sponsorde " + all_teams[idx] + ". Dit team heeft {:.2f} km gewandeld. Jij steunde dit team " \
        "€ {:.2f}  per kilometer. " \
        "Dat maakt dus € {:.2f}. Echter hebben we opgemerkt dat dit team zichzelf wat onderschat had " \
        "en dus meer wandelde dan hun verwachte {:.2f} km. U kan nog steeds het verwachte bedrag van " \
        "€ {:.2f} storten, we zijn dan ook content met elke cent. Eenmaal we dit ontvangen hebben maak je " \
        "ook effectief kans op een van onze mooie prijzen!  Op 15 mei zullen we de prijzen verdelen. " + "\n" + "\n" \
        "Alvast bedankt, \n\nDe leiding"
        message = message.format(all_actual[idx], all_amounts[idx],all_topay[idx], all_expected[idx],
                                             all_expectedtopay[idx])

    # Create the email to send
    full_email = ("From: {0} <{1}>\n"
                  "To: {2} <{3}>\n"
                  "Subject: {4}\n\n"
                  "{5}"
                  .format(your_name, your_email, name, email, subject, message).encode('utf-8'))

    # In the email field, you can add multiple other emails if you want
    # all of them to receive the same text
    try:
        server.sendmail(your_email, [email], full_email)
        print('Email to {} successfully sent!\n\n'.format(email))
    except Exception as e:
        print('Email to {} could not be sent :( because {}\n\n'.format(email, str(e)))

# Close the smtp server
server.close()