import os
os.system('pip install pandas openpyxl python-barcode pillow')

import pandas as pd
from barcode import Code128
from barcode.writer import ImageWriter
from google.colab import files

# Define the path where you want to save the barcodes in Colab's temporary filesystem
save_path = '/content/barcodes'

# Ensure the directory exists to save the barcodes
if not os.path.exists(save_path):
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
    print("Error: Column 'codes' not found in Excel file. Please ensure your Excel file has a column named 'codes'")
    exit()
except ValueError:
    print(f"Error: Sheet '{sheet_to_use}' not found in the Excel file")
    exit()
except Exception as e:
    print(f"Error reading Excel file: {str(e)}")
    exit()

# Function to create barcode from text
def create_barcode(text, filename):
    try:
        # Create barcode using Code128 format
        barcode = Code128(str(text), writer=ImageWriter())
        # Save the barcode
        barcode.save(os.path.join(save_path, filename))
    except Exception as e:
        print(f"Error creating barcode for {text}: {str(e)}")

# Loop through texts to create barcodes
for text in texts:
    # Clean the text for filename (removing invalid characters)
    filename = str(text).replace('@', 'at').replace('.', '_').replace('/', '_')
    create_barcode(text, filename)

print(f"Barcodes have been created in the directory: {save_path}")

# Optional: Download the generated barcodes
# This will zip the barcodes folder and prompt you to download it
!zip -r barcodes.zip {save_path}
files.download('barcodes.zip')
