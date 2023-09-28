import smtplib
from datetime import datetime
import os
import pandas as pd
from pprint import pprint
import time

# ---------- Gmail ----------
MY_EMAIL = os.environ.get("EMAIL")
PASSWORD = os.environ.get("PASS")

# ---------- CSV READ ----------
mail_df = pd.read_csv("mail_list.csv")
mail_dict = {row.nr: [row.nr, row.person, row.company, row.email, row.status] for (index, row) in mail_df.iterrows()}
# pprint(mail_dict)


def send_email(email, company, content):
    with smtplib.SMTP("smtp.gmail.com") as connection:
        connection.starttls()
        connection.login(user=MY_EMAIL, password=PASSWORD)
        connection.sendmail(from_addr=MY_EMAIL,
                            to_addrs=email,
                            msg=f"Subject: Cooperation with {company}!"
                                f"\n\n"
                                f"{content}"
                            )
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"Email sent to {company} on {timestamp}!")
    return f"Email sent {timestamp}"


def personalise_and_send(mailing_dictionary):
    for item in mailing_dictionary:
        with open("template.txt") as letter:
            content = letter.read()
            name = mailing_dictionary[item][1]
            company = mailing_dictionary[item][2]
            email = mailing_dictionary[item][3]
            status = mailing_dictionary[item][4]
            content_personalised = content.replace("[NAME]", name).replace("[COMPANY]", company)
            if status == "-":
                new_status = send_email(email, company, content_personalised)
                mailing_dictionary[item][4] = new_status
                time.sleep(5)
    return mailing_dictionary


timestamp_name = datetime.now().strftime("%Y-%m-%d--%H-%M")
updated_mail_dict = personalise_and_send(mail_dict)
updated_mail_df = pd.DataFrame(updated_mail_dict).transpose()
updated_mail_df.columns = ['nr', 'person', 'company', 'email', 'status']
updated_mail_df.to_csv(f'mail_list_{timestamp_name}.csv', mode="w", index=False, header=True, index_label='Index_name')
