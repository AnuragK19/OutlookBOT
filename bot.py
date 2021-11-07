import win32com.client
import datetime as dt
import pytz
import logging
import schedule
import time


def filterEmails():
    utc = pytz.UTC
    # basic config for logging
    logging.basicConfig(filename='logs.log', encoding='utf-8',
                        format='%(asctime)s %(levelname)-10s %(message)s', level=logging.INFO)

    logging.info(
        "Script started-------------[%s]-------------", dt.datetime.now())
    try:
        outlook = win32com.client.Dispatch(
            "Outlook.Application").GetNamespace("MAPI")
        logging.info(
            "Sucessfully initialized Outlook in a background process.")

        sourceFolderName = ''
        M_DestinationFN = ""
        N_DestinationFN = ""
        search_term = "Anurag"
        defaultSenderAddress = "anuraagkurmi@gmail.com"

        #sourceInbox = outlook.Folders[sourceFolderName].Folders['Inbox']
        sourceInbox = outlook.GetDefaultFolder(6)
        # Time interval for the mails to searched for
        today9am = dt.datetime.today().replace(
            hour=9, minute=00).strftime('%Y/%m/%d %H:%M %p')
        today9pm = dt.datetime.today().replace(
            hour=21, minute=00).strftime('%Y/%m/%d %H:%M %p')

        if(len(sourceFolderName) == 0 or len(M_DestinationFN) == 0 or len(N_DestinationFN) == 0 or len(search_term) == 0):
            logging.error("Some values are missing.")
            quit()
        else:
            logging.info("Config initialized successfully.")

        allMails = sourceInbox.Items.Restrict(
            "[SenderEmailAddress] = '" + defaultSenderAddress + "'")
        allMails.Sort("[ReceivedTime]", True)

        print("Total no of emails from the sender are : ", len(allMails))
        logging.info("Successfully fetched emails from specific sender")

        if len(allMails) > 0:
            for eachMail in allMails:
                if eachMail.UnRead == True and search_term in eachMail.Body.lower():
                    if eachMail.ReceivedTime >= utc.localize(dt.datetime.strptime(today9am, "%Y/%m/%d %H:%M %p")) and eachMail.ReceivedTime <= utc.localize(dt.datetime.strptime(today9pm, "%Y/%m/%d %H:%M %p")):
                        # eachMail.Read = True
                        # eachMail.Move(M_DestinationFN)
                        logging.info("[ %s ] Moved to [Not in shift]",
                                     eachMail.SenderEmailAddress)
                        print("[ %s ] Moved to [Not in shift]",
                              eachMail.SenderEmailAddress)
                    else:
                        # eachMail.Move(N_DestinationFN)
                        logging.info("[ %s ] Moved to [Ignore mail]",
                                     eachMail.SenderEmailAddress)
                        print("[ %s ] Moved to [Ignore mail]",
                              eachMail.SenderEmailAddress)
                else:
                    print("Rest of the mails are already read.")
                    logging.info(
                        "Script ended---------------[%s]-------------", dt.datetime.now())
                    break
        else:
            logging.info("No new mails found.")
            logging.info(
                "Script ended---------------[%s]-------------", dt.datetime.now())

    except Exception as ex:
        logging.error(
            " Already an instance of Outlook is running.Please close the existing instance of Outlook.", ex)
        logging.info(
            "Script ended---------------[%s]-------------", dt.datetime.now())
        quit()


# filterEmails("Inbox", "", "", "Rating..................Red",
#              "donotreply@mydomain.com")
schedule.every(1).minutes.do(filterEmails)

while 1:
    schedule.run_pending()
    time.sleep(1)
