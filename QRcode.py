import os
os.system('pip install qrcode[pil]')

import qrcode

# List of texts to convert to QR codes
texts = [
"0273399117.cp"
]

# Define the path where you want to save the QR codes
save_path = r'C:\Ashkan\Sunday'

# Ensure the directory exists to save the QR codes
if not os.path.exists(save_path):
    os.makedirs(save_path)

# Function to create QR code from text
def create_qr_code(text, filename):
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=10,
        border=4,
    )
    qr.add_data(text)
    qr.make(fit=True)

    img = qr.make_image(fill_color="black", back_color="white")
    img.save(os.path.join(save_path, f'{filename}.png'))

# Loop through texts to create QR codes with filenames matching the text
for text in texts:
    # Use the text itself for the filename, removing spaces if any, and replacing special characters
    filename = text.replace('@', 'at').replace('.', '_').replace('/', '_')
    create_qr_code(text, filename)

print(f"QR codes have been created in the directory: {save_path}")
