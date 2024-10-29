import os
import pandas as pd
from PIL import Image, ImageFont, ImageDraw

# Set the directory path of the script
dir_path = os.path.dirname(os.path.realpath(__file__))

# Load the Excel file
excel_path = f'{dir_path}/name_list.xlsx'