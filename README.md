# Certificate Generator

This project is a Python-based certificate generator that automatically creates personalized certificates for a list of participants provided in an Excel file. By using a customizable certificate template and participant data, this script can generate certificates as PDF files, which can be easily shared or printed.

## Objective

The goal of this project is to automate the process of creating certificates for participants, saving time and reducing manual effort. The project reads participant names from an Excel file, formats them appropriately, and adds them to a pre-designed certificate template.

## Project Structure


- **fonts/**: Stores the font file used to render the participant's name on the certificate.
- **template/**: Contains the certificate template image in PNG format.
- **output/**: The script will save generated certificates in this directory.
- **name_list.xlsx**: Excel file containing a list of participant names in a column titled "Name".
- **certificate_generator.py**: The main script that reads data, formats text, and generates the certificates.

## Workflow

1. **Load Excel Data**: The script reads names from `name_list.xlsx` in the "Name" column.
2. **Load Template and Fonts**: A certificate template (`certificate_template.png`) and a custom font (`CanvaSans-BoldItalic.ttf`) are loaded for text rendering.
3. **Generate Certificates**: For each participant:
   - The script centers the participant’s name on the template image.
   - The modified image is saved as a PDF file in the `output` directory.
4. **Save as PDF**: Each certificate is saved with the participant's name as the filename.

## Installation and Usage

### Prerequisites

Ensure that you have Python installed (preferably 3.x) and the following libraries:
- [Pandas](https://pandas.pydata.org/)
- [Pillow (PIL)](https://pillow.readthedocs.io/)

Install these dependencies by running:
```bash
pip install pandas pillow

```

## Setup

1. **Clone this repository**:
   ```bash
   git clone https://github.com/yourusername/certificate-generator.git
   cd certificate-generator
   ```

### Prepare Files

- Place the certificate template image in the `template/` directory.
- Add your custom font file to the `fonts/` directory.
- Create or update `name_list.xlsx` in the root directory with participant names under the "Name" column.
