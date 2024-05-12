import pywhatkit as kit
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials


scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name("./ticket.json", scope)
client = gspread.authorize(credentials)
sheet_url = 'https://docs.google.com/spreadsheets/d/1Yyli_6XocbKX2jyPUdpV1zhhjVxYNUd0ZXNAUvgHRGY/edit#gid=0'
sheet = client.open_by_url(sheet_url)
worksheet = sheet.get_worksheet(0)
expected_headers = ['Number','Sending Time', 'Status']
data2 = worksheet.get_all_records(expected_headers=expected_headers)
lenth= len(data2)

message = """প্রিয় ভাই,
KUET Muslim Community এর পক্ষ থেকে সবাইকে শুভেচ্ছা জানাচ্ছি। 
একটু ধৈর্য্য ধরে এসএমএসটি পড়ার জন্য অনুরোধ করা হচ্ছে।

২৩ মার্চ ২০২০, 
দিনটির তারিখ মনে না থাকলেও দিনটির কথা সবার মনে আছে বলে আমি ১০০% শিওর। কভিড -১৯ এর ভয়াবহতা সারা পৃথিবীতে ততক্ষনে ভয়াবহ ভাবে ছড়িয়ে পড়েছে। বাংলাদেশ এই দিনে ১ম লকডাউন এ যায়। মসজিদ থেকে আযানের পর এলান হতে সবাই ঘরে থাকুন ৫ জনের বেশি জামাতে শরিক হওয়া যাবে না। খেলার স্কোর বোর্ডের মত আতংকের সাথে সবাই কত মারা গেল তাই গণনা করতো।

এইরকম হাহাকার শোনা বা তা সহ্য করা যেমন কঠিন ছিল তেমনি একা একা এই পরিস্থিতিতে মানুষের পাশে দাঁড়ানো অনেক কঠিন ছিল।

শুর হয় এক সাথে কিছু করার প্রচেষ্টা। যেই চিন্তা সেই কাজ EX-KUET কিছু ভাই একত্রিত হয়ে শুরু হয় আলোচনা জন্ম নেয় "KUET MUSLIM COMMUNITY". এবং সেবছরই রমজানে প্রায় ৩ লক্ষ টাকার সাহায্য অসহায় মানুষের কাছে পৌঁছে দেবার সুযোগ দেন।

এই অসহায় মানুষের পাশে দাঁড়ানোর একটা সুযোগ সবার জন্য রাখতে চাই। Monthly donation এর মাধ্যমে আশা করি সবাই আমাদের পাশে থাকবেন।

মা 'আসসালাম ,
কুয়েট মুসলিম কমিউনিটি (KUETMC)"""

picture_path= "./KMC.png"
for index in range(lenth):
    try:
        row_index= index+2
        dataname= {}
        number = data2[index]["Number"]
        phone_number = f"+{number}"
        kit.sendwhats_image(phone_number, picture_path, message)
        sending_time = datetime.now().strftime("%d/%m/%Y : %H:%M")
        dataname["Sending Time"] = sending_time
        dataname["Status"] = "Sent"
        for column_name, value in dataname.items():
                        column_index = worksheet.find(column_name).col
                        worksheet.update_cell(row_index, column_index, value)
    except Exception as e:
        print(f"Error occurred while sending message to {phone_number}: {e}")
        sending_time = datetime.now().strftime("%d/%m/%Y : %H:%M")
        dataname["Sending Time"] = sending_time
        dataname["Status"] = "Sent"
        for column_name, value in dataname.items():
                        column_index = worksheet.find(column_name).col
                        worksheet.update_cell(row_index, column_index, value)
                        