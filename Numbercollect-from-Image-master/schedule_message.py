import pywhatkit as kit
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import schedule
import time

def send_message():
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_name("./ticket.json", scope)
    client = gspread.authorize(credentials)
    sheet_url = 'https://docs.google.com/spreadsheets/d/1Yyli_6XhocbKX2jyPUdpV1zhhjVxYNUd0ZXNAUvgHRGY/edit#gid=0'
    sheet = client.open_by_url(sheet_url)
    worksheet = sheet.get_worksheet(0)
    expected_headers = ['Number','Sending Time', 'Status']
    data2 = worksheet.get_all_records(expected_headers=expected_headers)
    length = len(data2)
    message = """
পদের নাম: QA/Test Automation/DevOps/Cyber Security Engineer

কোম্পানি: Quality UP Service (QUPS)
স্থানঃ খিলগাঁও, ঢাকা

সবাইকে Quality UP Service (QUPS) এর পক্ষ থেকে আন্তরিক শুভেচ্ছা ও অভিনন্দন। আপনারা যারা সফটওয়্যার কোয়ালিটি অস্সুরেন্স (SQA)/TA/DevOps/Cyber Security এ ক্যারিয়ার গড়তে ইচ্ছুক তাহলে এই জব পোস্টটি আপনার জন্য। প্রোগ্রামিং ল্যাঙ্গুয়েজ অর্থাৎ Python/Java/ JavaScript পারদর্শীদেরকে আমাদের টিমে নিতে চাই। একজন ইন্টার্ন হিসেবে আমাদের সফ্টওয়্যার পণ্যের গুণমান এবং নির্ভরযোগ্যতা নিশ্চিত করতে আমাদের QA Team এর সাথে একনিষ্ঠভাবে কাজ করতে হবে।

সুবিধা:
- লিভিং ফ্যাসিলিটিসহ আনুসাঙ্গিক সুবিধা।

শর্ত:
- ইন্টার্নশীপ এর পরে নির্দিষ্ট সময়ের জন্য জব সাপোর্ট দিতে হবে।
- অফিস স্পেসে (খিলগাঁও) রিলোকেশন করতে হবে।

আপনারা যারা প্রোগ্রামিং ল্যাঙ্গুয়েজ অর্থাৎ Python/Java/ JavaScript পারদর্শী এবং QA, Test Automation, DevOps এবং Cyber Security তে নিজের ক্যারিয়ার গড়তে আগ্রহী, তারা অনুগ্রহ করে নিম্নোক্ত গুগল ফর্মটি ফিলাপ করে আপনার জীবনবৃত্তান্ত জমা দিন।

Form Link: https://forms.gle/D78wiSmW1PMWMRB5A"""

    picture_path = "./1.png"
    for index in range(length):
        try:
            hour = datetime.now().hour
            minute = (datetime.now().minute + 5)
            row_index = index + 2
            dataname = {}
            number = data2[index]["Number"]
            phone_number = f"+{number}"
            kit.sendwhats_image(phone_number, picture_path, message, hour, minute)
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

# Schedule the job
schedule.every().day.at("09:17").do(send_message)  # Adjust time as needed

# Main loop
while True:
    schedule.run_pending()
    time.sleep(1)
