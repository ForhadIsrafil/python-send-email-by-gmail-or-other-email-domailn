import smtplib
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr

import pandas as pd

from CloudWatchLogger import logger
from Services.service import get_dataframe_from_db, get_worksheet
from Utils.config import TEMPLATE_EMAIL, MASTER_WITHDRAW_FILE_KEY, SENDER_EMAIL, SENDER_PASSWORD
from Utils.helpers import install_language, datetime_to_str

import warnings
warnings.filterwarnings('ignore')

install_language('Withdraw_Approved', 'locales')


def new_withdraw_emailer(user_email, amount_euro, request_day,
                         payperiod_start, payperiod_end, request_total_euro):
    try:
        sender_email = SENDER_EMAIL
        receiver_email = user_email
        bcc = 'daniel@wagenow.com'
        password = SENDER_PASSWORD

        # Create message container - the correct MIME type is multipart/alternative.
        message = MIMEMultipart('alternative')
        message['Subject'] = _("Your Withdraw is Approved and on it's way!")
        message['From'] = formataddr(('WageNow TM', 'wagenow@wagenow.com'))
        message['To'] = receiver_email
        message['Bcc'] = bcc

        email_recipients = bcc.split(",") + [receiver_email]

        # This example assumes the image is in the current directory
        fp = open('wagenow-logo_360.png', 'rb')
        msg_image = MIMEImage(fp.read())
        fp.close()

        # Define the image's ID as referenced above
        msg_image.add_header('Content-ID', '<image1>')

        text1 = _('Your Withdraw is approved and on it\'s way!')
        text2 = _('We\'ve got your back!')

        text3_variables = {
            'amountEURO1': amount_euro,
            'requestDay1': request_day
        }
        text3 = _('We received your request for %(amountEURO1)s on %(requestDay1)s.') % text3_variables

        text4 = _(
            'Depending on your bank, it may take up to one business day for the balance to be reflected on your account.')

        text5_variables = {
            'payperiodStart1': payperiod_start,
            'requestTotal1': request_total_euro
        }
        text5 = _(
            'For the pay period starting on %(payperiodStart1)s, your total approved requests (including this one) is %(requestTotal1)s.') % text5_variables

        text6 = _('Thank you for using WageNow, we\'re always here for you!')
        text7 = _('Have questions? Need assistance?')
        text8 = _('Ask us at ')

        email_link = '''<a href="mailto:wagenow@wagenow.com" rel="noopener" style="text-decoration: underline; color: #7569ff;" target="_blank" title="wagenow@wagenow.com">wagenow@wagenow.com</a>'''
        text9_variables = {
            'linkEmbed1': email_link
        }

        text_signature_name = _('Peter')
        text_signature_team = _('The WageNow Team')

        text9 = _(
            'The contents of this email message and any attachments are confidential and may be legally privileged. If you are not the intended recipient, notify the sender by replying to this email and delete this message and its attachments. To ensure delivery to your inbox, add %(linkEmbed1)s to your address book.') % text9_variables

        text10 = _(
            'DISCLOSURE: This communication is on behalf of WageNow Inc. ("WageNow"). This communication is not to be construed as legal, financial, accounting or tax advice and is for informational purposes only. This communication is not intended as a recommendation, offer or solicitation for the purchase or sale of any security. WageNow does not assume any liability for reliance on the information provided herein.')

        text11 = _('© 2023 WageNow Inc. All rights reserved.')

        # Create the body of the message (a plain-text and an HTML version).
        str_template_email = open(TEMPLATE_EMAIL, 'r').read()
        html = str_template_email.format(text_1=text1,
                                         text_2=text2,
                                         text_3=text3,
                                         text_4=text4,
                                         text_5=text5,
                                         text_6=text6,
                                         text_7=text7,
                                         text_8=text8,
                                         text_9=text9,
                                         text_10=text10,
                                         text_11=text11,
                                         text_signature_name=text_signature_name,
                                         text_signature_team=text_signature_team)

        part = MIMEText(html, 'html')
        message.attach(part)
        message.attach(msg_image)

        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender_email, password)
        server.sendmail(sender_email, email_recipients, message.as_string())
        server.quit()
    except Exception as e:
        logger.exception(e)


# users table
sql = '''
    SELECT 
        *
    FROM users as u
'''
users_data = get_dataframe_from_db(sql)

convert_cols = [
    'confirmed_at', 'last_login_at', 'start_of_period',
    'end_of_period', 'next_pay_date', 'created_at',
    'updated_at', 'deleted_at'
]

users_data = datetime_to_str(users_data, convert_cols)
users_data = users_data.fillna(0)

users_data = users_data[['id', 'start_of_period', 'end_of_period', 'email']]

# open master withdraws on shared drive
worksheet = get_worksheet(MASTER_WITHDRAW_FILE_KEY)
all_withdraws = pd.DataFrame(worksheet.get_all_records())
user_withdraws = all_withdraws.copy()
user_withdraws['createdAt'] = pd.to_datetime(user_withdraws['createdAt'])

# combine with start and end of pay period data
user_withdraws = user_withdraws.merge(users_data,
                                      right_on='id',
                                      left_on='userID',
                                      how='left')
user_withdraws = user_withdraws.drop('id', axis=1)

user_withdraws['start_of_period'] = pd.to_datetime(user_withdraws['start_of_period'])
user_withdraws['end_of_period'] = pd.to_datetime(user_withdraws['end_of_period'])
email_withdraws = user_withdraws.copy()
email_withdraws = user_withdraws[user_withdraws['emailSent'] != 'Yes']

for index, row in email_withdraws.iterrows():
    user_email = row['email']
    user_amount_euro = "\u20ac" + str(row['amount'])
    user_request_day = row['createdAt'].strftime("%d-%m-%Y")
    user_payperiod_start = row['start_of_period'].strftime("%d-%m-%Y")
    user_payperiod_end = row['end_of_period'].strftime("%d-%m-%Y")
    user_request_total_euro = "\u20ac" + str(
        user_withdraws.loc[(user_withdraws['createdAt'].dt.month == row['createdAt'].month)
                           & (user_withdraws['createdAt'].dt.year == row['createdAt'].year)
                           & (row['offsetMonth'] == 0)
                           & (user_withdraws['userID'] == row['userID']), 'amount'].sum())
    # send emails
    new_withdraw_emailer(user_email, user_amount_euro, user_request_day,
                         user_payperiod_start,
                         user_payperiod_end, user_request_total_euro)

    # update sheet
    r = all_withdraws.loc[all_withdraws['withdrawID'] == row['withdrawID']].index[0]
    worksheet.update_cell(r + 2, all_withdraws.columns.get_loc("emailSent") + 1, 'Yes')

logger.info("Executed")




# %%
