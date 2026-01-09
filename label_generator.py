import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import barcode
from barcode.writer import ImageWriter
import os

# ================== PATHS ==================
EXCEL_PATH = "70056721_Anthrilo Kids_8th Jan 26.xlsx"   # your file
OUTPUT_DIR = "generated_labels"
FONT_PATH = "DejaVuSans-Bold.ttf"  # put font in same folder or change path

os.makedirs(OUTPUT_DIR, exist_ok=True)

# ================== LOAD FILE ==================
if EXCEL_PATH.lower().endswith(".csv"):
    df = pd.read_csv(EXCEL_PATH)
else:
    df = pd.read_excel(EXCEL_PATH)

# ================== FIX COLUMN FORMAT ==================
df.columns = (
    df.columns
    .astype(str)
    .str.strip()
    .str.upper()
    .str.replace(" ", "_")
)

# ================== VALIDATE COLUMNS ==================
required_columns = {
    "PO_NO",
    "MODEL",
    "REF",
    "SIZE",
    "STYLE_CODE",
    "BARCODE",
    "TOTAL"
}

missing = required_columns - set(df.columns)
if missing:
    raise ValueError(f"Missing required columns: {missing}")

# ================== STATIC TEXT ==================
STATIC_TEXT = [
    "Distributor: Styli fze",
    "Importer: Retail Cart Trading Co.",
    "PO Box 86003 Riyadh, Kingdom of Saudi Arabia",
    "Importer: Landmark Online Retail LLC",
    "PO Box No. 25030 Dubai, UAE",
    "",
    "الموزع: ستايلي للتجارة ذ.م.م",
    "المستورد: شركة ريتيل كارت للتجارة",
    "ص.ب 86003 الرياض، المملكة العربية السعودية",
    "المستورد: لاندمارك أونلاين للتجزئة ذ.م.م",
    "ص.ب رقم 25030 دبي، الإمارات العربية المتحدة"
]

# ================== GENERATE LABELS ==================
label_counter = 0

for i, row in df.iterrows():
    # Get quantity from TOTAL column
    try:
        quantity = int(row['TOTAL'])
    except (ValueError, TypeError):
        print(f"Warning: Invalid quantity in row {i+1}, defaulting to 1")
        quantity = 1
    
    # Generate 'quantity' number of labels for this row
    for copy_num in range(quantity):
        label_counter += 1
        
        img = Image.new("RGB", (600, 900), "white")
        draw = ImageDraw.Draw(img)

        font_brand = ImageFont.truetype(FONT_PATH, 48)
        font_text = ImageFont.truetype(FONT_PATH, 26)
        font_small = ImageFont.truetype(FONT_PATH, 20)

        y = 20

        # Brand
        draw.text((240, y), "styli", font=font_brand, fill="black")
        y += 80

        # Dynamic fields
        fields = [
            f"Po No. : {row['PO_NO']}",
            f"Model : {row['MODEL']}",
            f"Ref. : {row['REF']}",
            f"Size. : {row['SIZE']}",
            f"Style code. : {row['STYLE_CODE']}",
        ]

        for line in fields:
            draw.text((50, y), line, font=font_text, fill="black")
            y += 35

        # Barcode
        barcode_value = str(row["BARCODE"]).strip()

        code128 = barcode.get("code128", barcode_value, writer=ImageWriter())
        barcode_path = os.path.join(OUTPUT_DIR, f"barcode_{label_counter}")
        code128.save(barcode_path)

        barcode_img = Image.open(barcode_path + ".png").resize((450, 120))
        img.paste(barcode_img, (75, y + 10))
        y += 150

        # Barcode number
        draw.text((200, y), barcode_value, font=font_text, fill="black")
        y += 40

        # Static text
        for line in STATIC_TEXT:
            draw.text((50, y), line, font=font_small, fill="black")
            y += 22

        # Footer
        y += 10
        draw.text((230, y), "Follow us", font=font_small, fill="black")
        y += 22
        draw.text((150, y), "@styliofficial      @styli_official", font=font_small, fill="black")

        # Save label
        output_path = os.path.join(OUTPUT_DIR, f"label_{label_counter}.png")
        img.save(output_path)

        print(f"Saved: {output_path} (Row {i+1}, Copy {copy_num+1}/{quantity})")

print("✅ ALL LABELS GENERATED SUCCESSFULLY")
