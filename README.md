# Barcode Label Generator

A simple and user-friendly tool to generate professional barcode labels for products. This application creates labels with product information, barcodes, and company details in both English and Arabic.

## What Does It Do?

This tool takes your product data from an Excel or CSV file and automatically creates barcode labels in PDF format. Each label includes:

- Company logo (Styli)
- Purchase Order number
- Product Model, Reference, Size, and Style Code
- Barcode (Code128 format)
- Distributor and Importer information in English and Arabic
- Social media handles

## Features

✨ **Easy to Use** - Simple graphical interface, no coding required  
📊 **Excel/CSV Support** - Works with your existing spreadsheets  
🏷️ **Professional Labels** - 50mm x 80mm labels at 300 DPI quality  
📄 **Single PDF Output** - All labels in one convenient PDF file  
🌐 **Bilingual** - Supports both English and Arabic text  
🎨 **Custom Branding** - Uses your company logo

## Installation

1. Make sure you have Python installed on your computer
2. Install the required packages:

```bash
pip install -r requirements.txt
```

## How to Use

1. **Run the Application**
   ```bash
   python label_generator_gui.py
   ```

2. **Select Your Excel/CSV File**
   - Click the "Browse" button next to "Excel/CSV File"
   - Choose your product data file
   - The file should contain these columns:
     - MODEL CODE
     - STYLI OPTION ID
     - SIZE
     - VENDOR STYLE NUMBER
     - STYLI SKU

3. **Choose Output Folder**
   - Click the "Browse" button next to "Output Folder"
   - Select where you want to save the PDF

4. **Generate Labels**
   - Click the "Generate Labels" button
   - Wait for the process to complete
   - Find your `labels.pdf` file in the output folder

## File Requirements

Your Excel or CSV file should have these column headers (the tool will automatically find them):
- **MODEL CODE** - Product model
- **STYLI OPTION ID** - Reference number
- **SIZE** - Product size
- **VENDOR STYLE NUMBER** - Style code
- **STYLI SKU** - Barcode number

The filename should start with the PO number, for example: `PO12345_products.xlsx`

## Label Specifications

- **Size**: 50mm × 80mm (2" × 3.15")
- **Resolution**: 300 DPI
- **Format**: PDF
- **Font**: Arial (Regular and Bold)
- **Barcode Type**: Code128

## Customization

To change the company logo:
1. Replace the `Stylilogo.jpeg` file with your own logo
2. Keep the same filename or update the code

## Requirements

- Python 3.6 or higher
- pandas - for reading Excel/CSV files
- Pillow (PIL) - for image processing
- python-barcode - for generating barcodes

## Troubleshooting

**"Could not find required columns"**  
Make sure your Excel file has the correct column headers listed above.

**"Logo not found"**  
Ensure `Stylilogo.jpeg` is in the same folder as the program.

**Font issues**  
The program uses Arial font. If not available, it will use a default font.

## Support

For issues or questions, please check that:
- Your Excel file has the correct column names
- All required packages are installed
- The logo file is present in the application folder

## License

Free to use for Styli company internal purposes.

---

Made with ❤️ for efficient label printing
