import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pprint
import random
import string
import datetime
import smtplib
import sys
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart

print('Connecting to Google Sheets....')
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name(
    'client_secret.json', scope)
client = gspread.authorize(creds)

sheet = client.open('Coupon').sheet1

pp = pprint.PrettyPrinter()


while True:
    mode = input('Enter 0 for managements, 1 for coupon redeeming\n')
    if mode == '0':
        while True:
            print("Updating coupons")
            coupon = sheet.get_all_values()
            managementMode = input(
                '0 to make all coupons valid\n1 to make all coupons invalid\n2 to generate coupons for new emails\n3 to force regenerate coupons for everyone\n4 to look at all coupons\nquit to quit\n')
            if managementMode == '0':
                # Select a range
                cell_list = sheet.range('C1:C'+str(len(coupon)))

                for cell in cell_list:
                    cell.value = '0'

                # Update in batch
                sheet.update_cells(cell_list)
                print('All coupons are now valid')

            if managementMode == '1':
                currentTimeString = str(datetime.datetime.now())
                # Select a range
                cell_list = sheet.range('C1:C'+str(len(coupon)))

                for cell in cell_list:
                    cell.value = currentTimeString

                # Update in batch
                sheet.update_cells(cell_list)
                print('All coupons are now invalid')

            if managementMode == '2':
                counter = 0
                for i in range(len(coupon)):
                    if coupon[i][1] == '':
                        counter += 1
                        sheet.update_cell(i+1, 2, str(i+1)+'.' +
                                        ''.join(random.choices(string.ascii_uppercase)))
                        sheet.update_cell(i+1, 3, '0')
                print('Added a total of '+str(counter)+' coupons')

            if managementMode == '3':
                counter = 0
                for i in range(len(coupon)):
                    sheet.update_cell(i+1, 2, str(i+1)+'.' +
                                    ''.join(random.choices(string.ascii_uppercase)))
                    sheet.update_cell(i+1, 3, '0')
                print('Changed a total of '+str(counter)+' coupons')

            if managementMode == '4':
                pp.pprint(coupon)

            if managementMode == '5':
                recievers = ["wenqi1016@gmail.com", "kh.kevinhuang.03@gmail.com"]
                msg = MIMEMultipart()
                msg['Subject'] = 'Coupon'
                msg['From'] = "lol233666@gmail.com"
                msg['To'] = ', '.join(recievers)
                msg.preamble = 'Our family reunion'
                s = smtplib.SMTP('localhost')
                s.sendmail("lol233666@gmail.com", recievers, msg.as_string())
                s.quit()



                # s = smtplib.SMTP('smtp.gmail.com', 587) 
                # s.starttls() 
                # s.login("sender_email_id", "sender_email_id_password") 
                # message = "Message_you_need_to_send"
                # s.sendmail("sender_email_id", "receiver_email_id", message) 
                # s.quit()

            if managementMode == 'quit':
                break

    if mode == '1':
        while True:
            print("Updating coupons list.....")
            coupon = sheet.get_all_values()
            coupon_code = input('\n\nEnter another coupon code to redeem it: ')

            if coupon_code == "quit":
                break

            coupon_code_array = coupon_code.split('.')
            try:
                if coupon[int(coupon_code_array[0])-1][1] == coupon_code and coupon[int(coupon_code_array[0])-1][2] == '0':
                    print('\nCoupon is valid!\nYou just used your coupon! Enjoy :)')
                    sheet.update_cell(
                        int(coupon_code_array[0]), 3, str(datetime.datetime.now()))
                else:
                    print('\n!!Coupon is invalid!!')
            except:
                print('\n!!Coupon is invalid!!')
                pass

    if mode == "quit":
        break
