import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import barcode
from barcode.writer import ImageWriter
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
import tempfile

try:
    import arabic_reshaper
    from bidi.algorithm import get_display
    ARABIC_SUPPORT = True
except ImportError:
    ARABIC_SUPPORT = False

class LabelGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Barcode Label Generator")
        self.root.geometry("600x400")
        self.root.configure(bg="#f0f0f0")
        
        self.excel_path = None
        self.output_dir = None
        
        # Title
        title_label = tk.Label(root, text="Barcode Label Generator", 
                              font=("Arial", 20, "bold"), bg="#f0f0f0", fg="#333")
        title_label.pack(pady=20)
        
        # Frame for file selection
        file_frame = tk.Frame(root, bg="#f0f0f0")
        file_frame.pack(pady=10, padx=20, fill="x")
        
        tk.Label(file_frame, text="Excel/CSV File:", font=("Arial", 11), 
                bg="#f0f0f0", width=15, anchor="w").grid(row=0, column=0, padx=5, pady=5)
        
        self.file_label = tk.Label(file_frame, text="No file selected", 
                                   font=("Arial", 10), bg="white", 
                                   relief="sunken", anchor="w")
        self.file_label.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        
        tk.Button(file_frame, text="Browse", command=self.browse_file,
                 bg="#4CAF50", fg="white", font=("Arial", 10, "bold"),
                 padx=10).grid(row=0, column=2, padx=5, pady=5)
        
        file_frame.columnconfigure(1, weight=1)
        
        # Frame for output directory
        output_frame = tk.Frame(root, bg="#f0f0f0")
        output_frame.pack(pady=10, padx=20, fill="x")
        
        tk.Label(output_frame, text="Output Folder:", font=("Arial", 11), 
                bg="#f0f0f0", width=15, anchor="w").grid(row=0, column=0, padx=5, pady=5)
        
        self.output_label = tk.Label(output_frame, text="No folder selected", 
                                     font=("Arial", 10), bg="white", 
                                     relief="sunken", anchor="w")
        self.output_label.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        
        tk.Button(output_frame, text="Browse", command=self.browse_output,
                 bg="#2196F3", fg="white", font=("Arial", 10, "bold"),
                 padx=10).grid(row=0, column=2, padx=5, pady=5)
        
        output_frame.columnconfigure(1, weight=1)
        
        # Progress bar
        self.progress = ttk.Progressbar(root, orient="horizontal", 
                                       length=400, mode="determinate")
        self.progress.pack(pady=20)
        
        # Status label
        self.status_label = tk.Label(root, text="Ready", font=("Arial", 10), 
                                    bg="#f0f0f0", fg="#666")
        self.status_label.pack(pady=5)
        
        # Generate button
        self.generate_btn = tk.Button(root, text="Generate Labels", 
                                     command=self.generate_labels,
                                     bg="#FF9800", fg="white", 
                                     font=("Arial", 14, "bold"),
                                     padx=30, pady=10, state="disabled")
        self.generate_btn.pack(pady=20)
        
    def browse_file(self):
        filename = filedialog.askopenfilename(
            title="Select Excel or CSV file",
            filetypes=[("Excel files", "*.xlsx *.xls"), 
                      ("CSV files", "*.csv"),
                      ("All files", "*.*")]
        )
        if filename:
            self.excel_path = filename
            self.file_label.config(text=Path(filename).name)
            self.check_ready()
    
    def browse_output(self):
        folder = filedialog.askdirectory(title="Select Output Folder")
        if folder:
            self.output_dir = folder
            self.output_label.config(text=folder)
            self.check_ready()
    
    def check_ready(self):
        if self.excel_path and self.output_dir:
            self.generate_btn.config(state="normal")
        else:
            self.generate_btn.config(state="disabled")
    
    def generate_labels(self):
        try:
            self.generate_btn.config(state="disabled")
            self.status_label.config(text="Loading data...", fg="#2196F3")
            self.root.update()
            
            # Extract PO number from filename (before underscore)
            filename = os.path.basename(self.excel_path)
            po_number = filename.split('_')[0] if '_' in filename else filename.split('.')[0]
            
            # Load data and find the correct header row
            # Column mapping: Excel column -> Our column name
            column_mapping = {
                'MODEL CODE': 'MODEL',
                'STYLI OPTION ID': 'REF',
                'SIZE': 'SIZE',
                'VENDOR STYLE NUMBER': 'STYLE_CODE',
                'STYLI SKU': 'BARCODE'
            }
            
            df = None
            header_row = None
            
            # Try to find the header row automatically
            for skip_rows in range(0, 20):  # Check first 20 rows
                try:
                    if self.excel_path.endswith('.csv'):
                        temp_df = pd.read_csv(self.excel_path, skiprows=skip_rows)
                    else:
                        temp_df = pd.read_excel(self.excel_path, skiprows=skip_rows)
                    
                    # Clean column names
                    temp_df.columns = temp_df.columns.astype(str).str.strip().str.upper()
                    
                    # Check if this row has the required columns (from Excel file)
                    if all(col in temp_df.columns for col in column_mapping.keys()):
                        df = temp_df
                        header_row = skip_rows
                        break
                except:
                    continue
            
            if df is None:
                messagebox.showerror("Error", 
                    f"Could not find required columns in the file.\n\n"
                    f"Required columns: {', '.join(column_mapping.keys())}\n\n"
                    f"Please make sure your Excel/CSV file contains these column names.")
                self.status_label.config(text="Error: Missing columns", fg="#f44336")
                self.generate_btn.config(state="normal")
                return
            
            # Rename columns to match our expected names
            df = df.rename(columns=column_mapping)
            
            # Add PO_NO column from filename
            df['PO_NO'] = po_number
            
            self.status_label.config(text=f"Found headers at row {header_row + 1}. Processing...", 
                                    fg="#2196F3")
            self.root.update()
            
            # Create output directory
            os.makedirs(self.output_dir, exist_ok=True)
            
            # Static text
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
            
            # Setup progress bar
            total_rows = len(df)
            self.progress["maximum"] = total_rows
            
            # Try to load a font (fallback to default if not found)
            try:
                font_brand = ImageFont.truetype("arialbd.ttf", 56)  # Bold font
                font_text = ImageFont.truetype("arialbd.ttf", 34)  # Bold for emphasis
                font_small = ImageFont.truetype("arialbd.ttf", 30)  # Bold and larger for static text
                font_small_en = ImageFont.truetype("arialbd.ttf", 22)  # Smaller bold for English
                font_footer = ImageFont.truetype("arial.ttf", 20)  # Compact footer
            except:
                try:
                    font_brand = ImageFont.truetype("C:\\Windows\\Fonts\\arialbd.ttf", 56)  # Bold font
                    font_text = ImageFont.truetype("C:\\Windows\\Fonts\\arialbd.ttf", 34)  # Bold
                    font_small = ImageFont.truetype("C:\\Windows\\Fonts\\arialbd.ttf", 30)  # Bold and larger
                    font_small_en = ImageFont.truetype("C:\\Windows\\Fonts\\arialbd.ttf", 22)  # Smaller bold for English
                    font_footer = ImageFont.truetype("C:\\Windows\\Fonts\\arial.ttf", 20)
                except:
                    font_brand = ImageFont.load_default()
                    font_text = ImageFont.load_default()
                    font_small = ImageFont.load_default()
                    font_small_en = ImageFont.load_default()
                    font_footer = ImageFont.load_default()
            
            # List to store all label images
            all_labels = []
            
            # Process each row
            for i, row in df.iterrows():
                self.status_label.config(text=f"Generating label {i+1} of {total_rows}...", 
                                        fg="#2196F3")
                self.progress["value"] = i + 1
                self.root.update()
                
                # 50.8mm x 81.3mm at 300 DPI
                img = Image.new("RGB", (600, 960), "white")
                draw = ImageDraw.Draw(img)
                
                y = 15  # Reduced from 20
                
                # Brand logo
                try:
                    logo_path = os.path.join(os.path.dirname(__file__), "Stylilogo.jpeg")
                    logo = Image.open(logo_path)
                    # Resize logo to fit nicely (adjust size as needed)
                    logo = logo.resize((250, 75), Image.Resampling.LANCZOS)
                    # Center the logo
                    logo_x = (600 - logo.width) // 2
                    img.paste(logo, (logo_x, y), logo if logo.mode == 'RGBA' else None)
                    y += 105  # Increased spacing for better separation
                except:
                    # Fallback to text if logo not found
                    draw.text((245, y), "styli", font=font_brand, fill="black")
                    y += 105
                
                # Dynamic fields
                fields = [
                    f"Po No. : {row['PO_NO']}",
                    f"Model : {row['MODEL']}",
                    f"Ref. : {row['REF']}",
                    f"Size. : {row['SIZE']}",
                    f"Style code. : {row['STYLE_CODE']}",
                ]
                
                for f in fields:
                    draw.text((300, y), f, font=font_text, fill="black", anchor="mm")
                    y += 43  # Slightly increased spacing
                
                # Barcode
                barcode_value = str(row["BARCODE"])
                code128 = barcode.get("code128", barcode_value, writer=ImageWriter())
                
                # Use temp directory for barcode
                with tempfile.TemporaryDirectory() as tmpdir:
                    barcode_path = os.path.join(tmpdir, f"bar_{i}")
                    code128.save(barcode_path)
                    
                    barcode_img = Image.open(barcode_path + ".png").resize((490, 140))
                    img.paste(barcode_img, (55, y + 10))
                
                y += 165  # Adjusted spacing
                
                # Static text
                for line in STATIC_TEXT:
                    # Right-align Arabic text, left-align English text
                    if line.strip() and any('\u0600' <= c <= '\u06FF' for c in line):
                        # Arabic text - right aligned
                        if ARABIC_SUPPORT:
                            # Reshape and apply bidi algorithm for proper Arabic rendering
                            reshaped_text = arabic_reshaper.reshape(line)
                            bidi_text = get_display(reshaped_text)
                            draw.text((550, y), bidi_text, font=font_small, fill="black", anchor="ra")
                        else:
                            # Fallback without reshaping
                            draw.text((550, y), line, font=font_small, fill="black", anchor="ra")
                    else:
                        # English text - left aligned with smaller font
                        draw.text((50, y), line, font=font_small_en, fill="black")
                    y += 32  # Adjusted spacing for larger bold font
                
                # Footer
                y += 13  # Adjusted spacing
                draw.text((230, y), "Follow us", font=font_footer, fill="black")
                y += 24  # Adjusted spacing
                draw.text((150, y), "@styliofficial      @styli_official", 
                         font=font_footer, fill="black")
                
                # Add to list
                all_labels.append(img.copy())
            
            # Save all labels to a single PDF
            if all_labels:
                pdf_path = os.path.join(self.output_dir, "labels.pdf")
                all_labels[0].save(pdf_path, save_all=True, append_images=all_labels[1:], resolution=300.0)
            
            # Success
            self.status_label.config(text=f"Success! Generated {total_rows} labels in PDF", 
                                    fg="#4CAF50")
            messagebox.showinfo("Success", 
                f"Successfully generated {total_rows} labels!\n\nSaved to: {os.path.join(self.output_dir, 'labels.pdf')}")
            
        except Exception as e:
            self.status_label.config(text=f"Error: {str(e)}", fg="#f44336")
            messagebox.showerror("Error", f"An error occurred:\n\n{str(e)}")
        
        finally:
            self.generate_btn.config(state="normal")
            self.progress["value"] = 0


if __name__ == "__main__":
    root = tk.Tk()
    app = LabelGeneratorApp(root)
    root.mainloop()
