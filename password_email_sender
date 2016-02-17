# Email sender that sends password to the mailing list
import smtplib
import logging
import logging.config
import yaml
import os
import pandas as pd
from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText
from copy import deepcopy


# Setup variables
config_directory=os.path.expanduser('~')+'/.user_config/'
email_subject='New password for bigpicture.lzd.co'

# Load logging settings
logging.config.dictConfig(yaml.load(open(os.path.join(config_directory,'pylogger.yaml'))))
logger=logging.getLogger('info_main_handler')

# Load central config file to get gmail username and API password
config=yaml.load(open(os.path.join(config_directory,'master_config.yaml')))
gmail_user=config['gmail']['username']
gmail_pwd=config['gmail']['password']
logger.info('Email credentials read')

# Main class
class info_emailer:
    """A generic class that allows sending of information though automated emails with simple mail merge
    functionality. Placeholder values are read through a CSV file, with the column names as placeholder names. There
    should be a column named "email" that has the receiver emails separated by semicolon"""

    def __init__(self, gmail_user, gmail_pwd):
        '''This class takes a .csv file as a placeholder input as well as a .txt file as a main email,
        and sends emails based on google mail smtp server by default'''

        self._gmail_user = gmail_user
        self._gmail_pwd = gmail_pwd

    def read_placeholders(self, placeholder_fn, directory='./'):
        '''Reads the placeholder and hence the receiver list'''

        self.placeholder_fullpath = os.path.join(directory, placeholder_fn)
        self.df_placeholders = pd.read_csv(self.placeholder_fullpath)
        logger.info('Placeholder .csv read successfully')

    def generate_email(self, subject, sender, receivers, message_text):
        '''Generate the email based on inputs'''

        msg =MIMEMultipart()
        msg['Subject'] = subject
        msg['From'] = sender
        msg['To'] = ','.join(receivers)
        msg_text = MIMEText(message_text)
        msg.attach(msg_text)

        logger.info('Email generated')
        return msg

    def send_email(self,sender,receivers,msg):
        mailServer = smtplib.SMTP('smtp.gmail.com', 587)
        #mailServer = smtplib.SMTP('localhost')
        mailServer.ehlo()
        mailServer.starttls()
        mailServer.ehlo()
        mailServer.login(self._gmail_user, self._gmail_pwd)
        mailServer.sendmail(sender,receivers,msg.as_string())
        #mailServer.sendmail()
        mailServer.quit()
        logger.info('email_sent')

    def mail_merge(self,email_subject,body_fn,body_directory='./'):
        """Use the CSV placeholders and do the mail merge"""

        self.message_original = open(os.path.join(body_directory, body_fn),'r').read()

        # Go through all rows in csv (sending multiple emails each time)
        for index, series in self.df_placeholders.iterrows():

            self.message = deepcopy(self.message_original) # Do this so it doesn't read the file again
            self.temp_dict = series[~series.index.isin(['email'])] # Exclude column email and all rest are placeholders

            # Go through each column (potential placeholders in the list)
            try:
                self.message = self.message.format(**self.temp_dict)
            except KeyError:
                logger.error('Columns mismatch between the csv file and the email body - there should only be 1 \
                             additional column called "email" in the csv that is the receivers list')

            email_list=str(series['email']).split(';') # Splits the semicolon delimited emails into a list
            self.composed=self.generate_email(email_subject,self._gmail_user,email_list,self.message)
            self.send_email(self._gmail_user,email_list,self.composed)


if __name__ == '__main__':

    sender = info_emailer(gmail_user,gmail_pwd)
    sender.read_placeholders('bpr_email_list_adhoc.csv')
    sender.mail_merge(email_subject,'message_body.txt')