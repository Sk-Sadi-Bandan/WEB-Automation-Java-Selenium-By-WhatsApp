import os
import re
import pytesseract
from PIL import Image
import phonenumbers
# Directory containing the images
directory = './picture/'

# List to store image names
image_names = []

# Iterate over all files in the directory
for filename in os.listdir(directory):
    # Check if the file ends with '.jpeg' (case-insensitive)
    if filename.lower().endswith('.jpeg'):
        # Add the image name to the list
        image_names.append(filename)
        
for filename in image_names:
    # Open the image file
    image = Image.open(f'./picture/{filename}')

    text = pytesseract.image_to_string(image)

    country_codes = "AF"#, "AL", "DZ", "AS", "AD", "AO", "AI", "AQ", "AG", "AR", "AM", "AW", "AC", "AU", "AT", "AZ", "BS", "BH", "BD", "BB", "BY", "BE", "BZ", "BJ", "BM", "BT", "BO", "BA", "BW", "BR", "VG", "BN", "BG", "BF", "MM", "BI", "KH", "CM", "CA", "CV", "KY", "CF", "TD", "CL", "CN", "CX", "CC", "CO", "KM", "CG", "CK", "CR", "HR", "CU", "CY", "CZ", "CD", "DK", "DG", "DJ", "DM", "DO", "EC", "EG", "SV", "GQ", "ER", "EE", "ET", "FK", "FO", "FJ", "FI", "FR", "GF", "PF", "GA", "GM", "GE", "DE", "GH", "GI", "GR", "GL", "GD", "GP", "GU", "GT", "GN", "GW", "GY", "HT", "VA", "HN", "HK", "HU", "IS", "IN", "ID", "IR", "IQ", "IE", "IM", "IL", "IT", "CI", "JM", "JP", "JE", "JO", "KZ", "KE", "KI", "KW", "KG", "LA", "LV", "LB", "LS", "LR", "LY", "LI", "LT", "LU", "MO", "MK", "MG", "MW", "MY", "MV", "ML", "MT", "MH", "MQ", "MR", "MU", "YT", "MX", "FM", "MD", "MC", "MN", "ME", "MS", "MA", "MZ", "NA", "NR", "NP", "NL", "AN", "NC", "NZ", "NI", "NE", "NG", "NU", "NF", "KP", "MP", "NO", "OM", "PK", "PW", "PS", "PA", "PG", "PY", "PE", "PH", "PN", "PL", "PT", "PR", "QA", "CG", "RE", "RO", "RU", "RW", "BL", "SH", "KN", "LC", "MF", "PM", "VC", "WS", "SM", "ST", "SA", "SN", "RS", "SC", "SL", "SG", "SX", "SK", "SI", "SB", "SO", "ZA", "KR", "SS", "ES", "LK", "SD", "SR", "SJ", "SZ", "SE", "CH", "SY", "TW", "TJ", "TZ", "TH", "TL", "TG", "TK", "TO", "TT", "TN", "TR", "TM", "TC", "TV", "UG", "UA", "AE", "GB", "US", "UY", "VI", "UZ", "VU", "VE", "VN", "WF", "EH", "YE", "ZM", "ZW"]

# Iterate over each country code
    # for country_code in country_codes:

    matches = phonenumbers.PhoneNumberMatcher(text, country_codes)
    # print(matches)

# Print the extracted phone numbers
    for match in matches:
        phone=phonenumbers.format_number(match.number, phonenumbers.PhoneNumberFormat.E164)
        with open('phone_numbers.txt', 'a', encoding='utf-8') as file:
                file.write(phone + '\n')
            
            # number = match.number
            # parsed_number = phonenumbers.parse(number, "XX")
            # print(parsed_number)
            # country = geocoder.description_for_number(parsed_number, "en")
            # print(f"Extracted phone number in {country_code}: {number} ({country})")
            