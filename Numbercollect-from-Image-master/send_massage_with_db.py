import sqlite3
import pywhatkit as kit
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os
import time
# Connect to the SQLite database
connection = sqlite3.connect('messages.db')
cursor = connection.cursor()

# Create a table to store the messages
cursor.execute('''CREATE TABLE IF NOT EXISTS messages (
                    id INTEGER PRIMARY KEY,
                    number TEXT,
                    message TEXT
                )''')




# Function to check if the message is the same for a given number
def is_message_same(number, message):
    cursor.execute("SELECT * FROM messages WHERE number = ? AND message = ?", (number, message))
    result = cursor.fetchone()
    return result is not None

# Function to insert a new message record into the database
def insert_message(number, message):
    cursor.execute("INSERT INTO messages (number, message) VALUES (?, ?)", (number, message))
    connection.commit()

# Your existing code
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name("./ticket.json", scope)
client = gspread.authorize(credentials)
sheet_url = 'https://docs.google.com/spreadsheets/d/1Yyli_6XocbKX2jyPUdpV1zhhjVxYNUd0ZXNAUvgHRGY/edit#gid=0'
sheet = client.open_by_url(sheet_url)
worksheet = sheet.get_worksheet(0)
expected_headers = ['Number','Sending Time', 'Status']
data2 = worksheet.get_all_records(expected_headers=expected_headers)
length = len(data2)

message = """"প্রিয় ভাই,  
KUET Muslim Community এর পক্ষ থেকে সবাইকে শুভেচ্ছা জানাচ্ছি।

রমাদানের ইফতার অনুষ্ঠানের পরামর্শ অনুযায়ী আমরা আবারো ঈদের পরে এপ্রিলের ২৭ তারিখ KUET Muslim Community এর পক্ষ থেকে একটি গেট টুগেদার এর আয়োজন করেছি। এ সময়ে KUET Muslim Community এর কে কোন প্রজেক্টে অংশগ্রহণ করব এবং আগামী রমজান পর্যন্ত আমাদের কার্যক্রম কি হবে সে বিষয়ে Road map নিয়ে আলোচনা হবে, ইনশাআল্লাহ।
আপনার উপস্থিতি আমাদের কাছে একান্ত কাম্য।

তারিখঃ   ২৭ এপ্রিল ২০২৪, রোজঃ শনিবার
ঠিকানাঃ  House 32, Road No. 25, Dhaka 1216, Lift 3, Imperial Diba Tower (পল্লবী ঝিল জামে মসজিদের কাছে,  মিরপুর-১২, ঢাকা)
Google Map : https://maps.app.goo.gl/zeLJaeUFd9fr8oaz5

সময়সূচী  (সম্ভাব্য)ঃ
বিকাল ৪টা থেকে রাত ৯টা পর্যন্ত আলোচনা ও পরামর্শ।

রেজিস্ট্রেশন লিংক : https://docs.google.com/forms/d/e/1FAIpQLScGFCx--StfFqijDRkdWQS12HXtMtlpt-ARb7mw_qKaLlew-A/viewform

মা 'আসসালাম ,
কুয়েট মুসলিম কমিউনিটি (KUETMC)

KUET Muslim Community
WhatsApp Group Link: https://chat.whatsapp.com/JRxg4aP1yXeDXGG3oZpDWa"""

picture_path = "./KMC Get Together.png"
# Shutdown
hours = 3.5
shutdown_time = time.time() + hours * 3600 

for index in range(length):


    try:

        # Shutdown
        if time.time() >= shutdown_time:
         os.system("shutdown /s /t 1")


        row_index = index + 2
        dataname = {}
        number = data2[index]["Number"]
        phone_number = f"+{number}"
        
        # Check if the message is the same for this number
        if not is_message_same(number, message):
            # Send the message
            kit.sendwhats_image(phone_number, picture_path, message, datetime.now().hour, (datetime.now().minute + 1))
            
            # Store the message in the database
            insert_message(number, message)
            
            # Update the sending time and status in the Google Sheet
            sending_time = datetime.now().strftime("%d/%m/%Y : %H:%M")
            dataname["Sending Time"] = sending_time
            dataname["Status"] = "Sent"
            
            for column_name, value in dataname.items():
                column_index = worksheet.find(column_name).col
                worksheet.update_cell(row_index, column_index, value)
    except Exception as e:
        print(f"Error occurred while sending message to {phone_number}: {e}")
        # Update the sending time and status in the Google Sheet
        sending_time = datetime.now().strftime("%d/%m/%Y : %H:%M")
        dataname["Sending Time"] = sending_time
        dataname["Status"] = "Failed"
        for column_name, value in dataname.items():
            column_index = worksheet.find(column_name).col
            worksheet.update_cell(row_index, column_index, value)

# Close the database connection
connection.close()
