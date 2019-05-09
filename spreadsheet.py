import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pprint
import random
import string

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json',scope)
client = gspread.authorize(creds)

sheet = client.open('Coupon').sheet1

pp=pprint.PrettyPrinter()
coupon = sheet.get_all_values()
coupon_code = input('Enter a coupon code: ')
coupon_code = coupon_code.split('.')

# pp.pprint(coupon)
for i in range(len(coupon)):
    sheet.update_cell(i+1,2,str(i+1)+'.'+''.join(random.choices(string.ascii_uppercase + string.digits, k=2)))

pp.pprint(coupon)