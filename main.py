import pandas as pd
import datetime
import smtplib  # Simple mail transfer protocol Library

GMAIL_ID = 'vritikamalhotra23@gmail.com'
GMAIL_PSWD = 'xxxxxxxx'


def sendEmail(to, sub, msg):
    print("Email to {} sent with subject {} and message {}".format(to, sub, msg))
    s = smtplib.SMTP('smtp.gmail.com', 587)  # SMTP infor for gmail.

    # Generating SMTP session
    s.starttls()
    s.login(GMAIL_ID, GMAIL_PSWD)
    s.sendmail(GMAIL_ID, to, f"Subject : {sub}\n\n{msg}")
    s.quit()


if __name__ == '__main__':

    df = pd.read_excel("F:\\PythonScripts\\pycharm\\BirthdayWisher\\data.xlsx")
    today = datetime.datetime.now().strftime("%d-%m")  # Coverts datetime into string format.
    yearNow = datetime.datetime.now().strftime("%Y")
    writeInd = []
    for index, item in df.iterrows():
        bday = item['Birthday'].strftime("%d-%m")
        if ((today == bday) and (yearNow not in str(item['Year']))):
            sendEmail(item['Email'], "Happy Birthday From Vritika Malhotra", item['Dialogue'])
            writeInd.append(index)
    if (len(writeInd) != 0):
        for i in writeInd:
            yr = df.loc[i, 'Year']
            df.loc[i, 'Year'] = str(yr) + ',' + str(yearNow)
        df.to_excel('F:\\PythonScripts\\pycharm\\BirthdayWisher\\data.xlsx', index=False)
