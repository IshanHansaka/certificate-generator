import os
import pandas as pd
from PIL import Image, ImageFont, ImageDraw

# Set the directory path of the script
dir_path = os.path.dirname(os.path.realpath(__file__))

# Load the Excel file
excel_path = f'{dir_path}/name_list.xlsx'

# Load the Excel file and handle any errors
try:
    df = pd.read_excel(excel_path)
    print("Excel file loaded successfully.")
except Exception as e:
    print(f"Error in loading the Excel file: {e}")
    exit()

# Check for required columns
if "Name" not in df.columns:
    print("Error: Required 'Name' column not found in the Excel file.")
    exit()

# Extract data from the Excel columns
nameList = df["Name"].fillna("Unknown Name")  # Handle missing names
length = len(df) # Get exact row count

# Configure fonts for name and team name
name_font_path = f'{dir_path}/fonts/CanvaSans-BoldItalic.ttf'

font_color = "#20443b" # Add font color
font_size = 60 # Add font size

# Font loading
try:
    name_font = ImageFont.truetype(name_font_path, font_size)
    print("Fonts loaded successfully.")
except Exception as e:
    print(f"Error loading fonts: {e}")
    exit()

template_width = 2000 # Add template width in px
template_height = 1414 # Add template height in px

# Create an output directory to save generated certificates
output_folder = f'{dir_path}/output'
os.makedirs(output_folder, exist_ok=True)

# Generate certificates
for i in range(length):
    # Load the certificate template
    template_path = f'{dir_path}/template/certificate_template.png'
    
    try:
        template_image = Image.open(template_path)
        print("Template image loaded successfully.")
    except Exception as e:
        print(f"Error loading template image: {e}")
        continue

    # Extract name and team for each participant, handle missing values
    try:
        name_of_participant = str(nameList[i]).strip().title()
    except Exception as e:
        print(f"Error processing participant data at row {i + 1}: {e}")
        continue
    
    # Prepare the template image to be edited
    editable_image = ImageDraw.Draw(template_image)
    
    name_text_box = editable_image.textbbox((0, 0), name_of_participant, font=name_font)
    name_text_width = name_text_box[2] - name_text_box[0]
    name_text_height = name_text_box[3] - name_text_box[1]

    # Adjust the y-position offset (positive to move down, negative to move up)
    y_offset = -25  # Change this value as needed

    # Calculate centered X and adjusted Y position for the name
    name_x_position = (template_width - name_text_width) // 2
    name_y_position = ((template_height - name_text_height) // 2) + y_offset  # Adjusted Y position

    # Add name to the certificate
    editable_image.text((name_x_position, name_y_position), name_of_participant, fill=font_color, font=name_font)
    
    # Create a filename by sanitizing the participant's name
    pdf_filename = f"{name_of_participant.replace(' ', '_')}.pdf"
    output_path = os.path.join(output_folder, pdf_filename)
    
    try:
        # Save the customized certificate as a PDF file
        template_image.save(output_path, "PDF", resolution=100.0)
        print(f"Certificate created for {name_of_participant} at {output_path}")
    except Exception as e:
        print(f"Error saving certificate for {name_of_participant}: {e}")