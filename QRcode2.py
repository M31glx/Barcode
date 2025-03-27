import os
os.system('pip install pandas openpyxl qrcode[pil]')

import pandas as pd
import qrcode
from google.colab import files
import shutil

# Define the path where you want to save the QR codes
save_path = '/content/qrcodes'

# Clear previous QR codes cache (optional)
if os.path.exists(save_path):
    shutil.rmtree(save_path)
    print(f"Cleared previous cache at {save_path}")
os.makedirs(save_path)

# Upload your Excel file
print("Please upload your Excel file:")
uploaded = files.upload()

# Get the filename of the uploaded file
file_name = list(uploaded.keys())[0]

# Specify the sheet name
sheet_to_use = "Sheet1"

# Read the Excel file
try:
    df = pd.read_excel(file_name, sheet_name=sheet_to_use)
    # Assuming the column with the data is named 'codes'
    texts = df['codes'].tolist()
except KeyError:
    print("Error: Column 'codes' not found in Excel file.")
    exit()
except ValueError:
    print(f"Error: Sheet '{sheet_to_use}' not found in the Excel file")
    exit()
except Exception as e:
    print(f"Error reading Excel file: {str(e)}")
    exit()

# Function to create QR code from text
def create_qr_code(text, filename):
    try:
        qr = qrcode.QRCode(
            version=1,  # Size of the QR code (1-40, auto-adjusted if fit=True)
            error_correction=qrcode.constants.ERROR_CORRECT_L,  # Low error correction (7%)
            box_size=10,  # Size of each box in pixels
            border=4,  # Border thickness in boxes
        )
        qr.add_data(text)
        qr.make(fit=True)  # Automatically adjust version to fit data
        
        img = qr.make_image(fill_color="black", back_color="white")
        img.save(os.path.join(save_path, f"{filename}.png"))
    except Exception as e:
        print(f"Error creating QR code for {text}: {str(e)}")

# Loop through texts to create QR codes
for text in texts:
    # Clean the text for filename (removing invalid characters)
    filename = str(text).replace('@', 'at').replace('.', '_').replace('/', '_')
    create_qr_code(text, filename)

print(f"QR codes have been created in the directory: {save_path}")

# Zip and download QR codes
!zip -r qrcodes.zip {save_path}
files.download('qrcodes.zip')
