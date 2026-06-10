import pandas as pd
from reportlab.lib.pagesizes import inch, letter
from reportlab.pdfgen import canvas
import tkinter as tk
from tkinter import messagebox, ttk, filedialog
import win32print
import os
import time
import traceback
import json
from barcode import UPCA
from barcode.writer import ImageWriter
from reportlab.lib.utils import ImageReader

# Config file to store last used paths
CONFIG_FILE = os.path.join(os.path.expanduser("~"), "label_printer_config.json")

def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r') as f:
                return json.load(f)
        except:
            return {"data_path": ""}
    return {"data_path": ""}

def save_config(config):
    try:
        with open(CONFIG_FILE, 'w') as f:
            json.dump(config, f)
    except Exception as e:
        print(f"Warning: Could not save config: {e}")

def find_data_folder():
    """
    Find the data folder. First checks saved location, then prompts user.
    This works across all systems - no hardcoded drives.
    """
    config = load_config()
    
    # Try last used path first
    if config.get("data_path") and os.path.exists(config["data_path"]):
        return config["data_path"]
    
    # Prompt user to select folder
    messagebox.showinfo(
        "Select Data Folder", 
        "Please locate the PRINT_LABELS_BOXES_TINS folder.\n\n"
        "This is typically on a network drive or shared location."
    )
    
    path = filedialog.askdirectory(
        title="Select PRINT_LABELS_BOXES_TINS folder",
        mustexist=True
    )
    
    if path:
        # Verify this looks like the correct folder
        excel_file = os.path.join(path, "UPC_With_Barcodes_FOR_PRINT.xlsx")
        if not os.path.exists(excel_file):
            retry = messagebox.askyesno(
                "Confirmation", 
                f"The file 'UPC_With_Barcodes_FOR_PRINT.xlsx' was not found in this folder.\n\n"
                f"Selected: {path}\n\n"
                f"Do you want to use this folder anyway?"
            )
            if not retry:
                return find_data_folder()  # Try again
        
        config["data_path"] = path
        save_config(config)
        return path
    else:
        raise ValueError("Data folder not selected. Application cannot continue.")

def initialize_app():
    """Initialize application with data folder and files"""
    data_folder = find_data_folder()
    data_file = os.path.join(data_folder, "UPC_With_Barcodes_FOR_PRINT.xlsx")
    label_dir = os.path.join(data_folder, "Labels")
    barcode_dir = os.path.join(data_folder, "Barcodes")
    
    # Create subdirectories if they don't exist
    try:
        os.makedirs(label_dir, exist_ok=True)
        os.makedirs(barcode_dir, exist_ok=True)
    except Exception as e:
        messagebox.showwarning(
            "Directory Creation", 
            f"Could not create subdirectories:\n{e}\n\nFiles will be saved to main folder."
        )
        label_dir = data_folder
        barcode_dir = data_folder
    
    # Load Excel file
    try:
        if not os.path.exists(data_file):
            raise FileNotFoundError(
                f"Excel file not found:\n{data_file}\n\n"
                f"Please ensure 'UPC_With_Barcodes_FOR_PRINT.xlsx' exists in the data folder."
            )
        
        df = pd.read_excel(data_file, converters={'Tin SKU': str})
        
        # Validate required columns
        required_columns = ["Tea", "Tao ID", "Tin SKU", "Barcode_UPC"]
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            raise ValueError(f"Excel file is missing required columns: {', '.join(missing_columns)}")
        
        products = df["Tea"].tolist()
        
        if not products:
            raise ValueError("No products found in Excel file.")
        
        return data_file, label_dir, barcode_dir, df, products
        
    except Exception as e:
        error_msg = f"Error reading Excel file:\n{str(e)}\n\nPlease check the file and try again."
        print(f"Full error:\n{traceback.format_exc()}")
        messagebox.showerror("Error", error_msg)
        raise

def split_text(c, text, max_width, font_name, font_size):
    """Split text into multiple lines to fit within max_width"""
    c.setFont(font_name, font_size)
    lines = []
    current_line = ""
    for word in text.split():
        test_line = f"{current_line} {word}".strip()
        if c.stringWidth(test_line, font_name, font_size) <= max_width:
            current_line = test_line
        else:
            if current_line:
                lines.append(current_line)
            current_line = word
    if current_line:
        lines.append(current_line)
    return lines

def generate_label(tea, quantity, printer_name, df, label_dir, barcode_dir):
    """Generate and print labels for selected tea product"""
    try:
        # Get product data
        product = df[df["Tea"] == tea].iloc[0]
        tao_id = product["Tao ID"]
        sku_value = product["Tin SKU"]
        sku = "N/A" if pd.isna(sku_value) else str(sku_value)
        barcode_value = product["Barcode_UPC"]
        
        if pd.isna(barcode_value):
            raise ValueError(f"Barcode is missing for product: {tea}")
        
        barcode = str(int(barcode_value))
        
        # Create safe filename
        safe_tea = "".join(c for c in tea if c.isalnum() or c in (' ', '_')).rstrip()
        output_file = os.path.join(label_dir, f"label_{safe_tea}.pdf")
        barcode_path = os.path.join(barcode_dir, f"{safe_tea}_barcode")
        
        # Generate barcode image
        barcode_writer = ImageWriter()
        barcode_file = UPCA(barcode, writer=barcode_writer)
        barcode_file.save(barcode_path)
        
        # Create PDF with labels
        c = canvas.Canvas(output_file, pagesize=letter)
        label_width, label_height = 2 * inch, 2 * inch
        labels_per_row, labels_per_col = 4, 5
        labels_per_page = labels_per_row * labels_per_col
        x_margin = 0.161 * inch
        y_margin = 0.321 * inch
        x_gap = 0.011 * inch
        y_gap = 0.1 * inch
        x_shift = -0.039 * inch
        y_shift = -0.098 * inch
        
        for i in range(quantity):
            if i > 0 and i % labels_per_page == 0:
                c.showPage()
            
            row = (i % labels_per_page) // labels_per_row
            col = (i % labels_per_page) % labels_per_row
            x = x_margin + col * (label_width + x_gap) + x_shift
            y = letter[1] - y_margin - (row + 1) * label_height - row * y_gap
            
            # TAO ID - right aligned, both parts same size
            label_text = "TAO ID: "
            tao_id_y = y + label_height - 0.24 * inch + y_shift
            c.setFont("Times-Bold", 16)
            
            label_width_text = c.stringWidth(label_text, "Times-Bold", 16)
            id_width = c.stringWidth(str(tao_id), "Times-Bold", 16)
            total_width = label_width_text + id_width
            right_margin = x + label_width - 0.16 * inch
            
            if total_width > (label_width - 0.32 * inch):
                label_x = right_margin - total_width
            else:
                label_x = right_margin - id_width - label_width_text
            
            c.drawString(label_x, tao_id_y, label_text)
            c.drawString(label_x + label_width_text, tao_id_y, str(tao_id))
            
            # Tea name (with text wrapping)
            c.setFont("Times-Bold", 10)
            tea_text = f"TEA: {tea}"
            max_width = label_width - 0.32 * inch
            tea_lines = split_text(c, tea_text, max_width, "Times-Bold", 10)
            tea_y = y + 1.52 * inch + y_shift
            line_height = 0.16 * inch
            for line in tea_lines[:2]:
                c.drawString(x + 0.16 * inch, tea_y, line)
                tea_y -= line_height
            
            # SKU
            sku_y = y + 1.28 * inch + y_shift if len(tea_lines) == 1 else y + 1.12 * inch + y_shift
            c.drawString(x + 0.16 * inch, sku_y, f"SKU: {sku}")
            
            # QTY field (empty for manual entry)
            qty_y = y + 1.04 * inch + y_shift if len(tea_lines) == 1 else y + 0.88 * inch + y_shift
            c.drawString(x + 0.16 * inch, qty_y, "QTY: ")
            
            # Barcode image
            barcode_img = ImageReader(f"{barcode_path}.png")
            c.drawImage(barcode_img, x + 0.16 * inch, y + 0.16 * inch + y_shift, 
                       width=1.76 * inch, height=0.56 * inch)
            
            # Label border
            c.rect(x + 0.08 * inch, y + 0.08 * inch, 
                  label_width - 0.16 * inch, label_height - 0.16 * inch, 
                  stroke=1, fill=0)
        
        c.save()
        print(f"PDF saved to: {output_file}")
        
        # Clean up barcode image
        try:
            os.remove(f"{barcode_path}.png")
        except:
            pass
        
        time.sleep(0.5)
        
        # Print the PDF
        try:
            win32print.SetDefaultPrinter(printer_name)
            
            import subprocess
            import sys
            
            # Try SumatraPDF first (best for silent printing)
            sumatra_paths = [
                r"C:\Program Files\SumatraPDF\SumatraPDF.exe",
                r"C:\Program Files (x86)\SumatraPDF\SumatraPDF.exe",
                os.path.join(os.path.dirname(sys.executable), "SumatraPDF", "SumatraPDF.exe")
            ]
            
            printed = False
            for path in sumatra_paths:
                if os.path.exists(path):
                    print(f"Printing with SumatraPDF: {path}")
                    subprocess.call([path, "-print-to", printer_name, "-silent", output_file])
                    printed = True
                    break
            
            # Try Adobe Reader if SumatraPDF not available
            if not printed:
                adobe_paths = [
                    r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
                    r"C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe",
                    r"C:\Program Files\Adobe\Reader 11.0\Reader\AcroRd32.exe",
                    r"C:\Program Files (x86)\Adobe\Reader 11.0\Reader\AcroRd32.exe"
                ]
                
                for path in adobe_paths:
                    if os.path.exists(path):
                        print(f"Printing with Adobe: {path}")
                        subprocess.call([path, "/t", output_file, printer_name])
                        printed = True
                        break
            
            # Fallback to Windows default print
            if not printed:
                print("Using Windows default print")
                if sys.platform == 'win32':
                    os.startfile(output_file, "print")
            
            messagebox.showinfo("Success", 
                f"Printing {quantity} label(s) for {tea}\n"
                f"Printer: {printer_name}\n\n"
                f"PDF saved to:\n{output_file}"
            )
        
        except Exception as print_error:
            error_msg = f"Error printing: {str(print_error)}"
            print(f"Print error details:\n{traceback.format_exc()}")
            messagebox.showerror("Print Error", 
                f"{error_msg}\n\n"
                f"The PDF was created successfully at:\n{output_file}\n\n"
                f"You can print it manually."
            )
    
    except Exception as e:
        error_msg = f"Error generating label: {str(e)}"
        print(f"Full error:\n{traceback.format_exc()}")
        messagebox.showerror("Error", error_msg)

# Main application
try:
    DATA_FILE, LABEL_DIR, BARCODE_DIR, df, products = initialize_app()
    
    root = tk.Tk()
    root.title("The Tao of Tea - Label Printer")
    root.geometry("650x600")
    root.configure(bg="#f5f1e8")  # Warm cream background
    
    # Header Frame with Company Name
    header_frame = tk.Frame(root, bg="#2d5016", height=80)  # Deep tea green
    header_frame.pack(fill=tk.X, pady=0)
    header_frame.pack_propagate(False)
    
    company_label = tk.Label(
        header_frame, 
        text="The Tao of Tea",
        font=("Georgia", 28, "bold"),
        fg="#d4af37",  # Gold color
        bg="#2d5016"
    )
    company_label.pack(pady=20)
    
    subtitle_label = tk.Label(
        header_frame,
        text="Label Printing System",
        font=("Arial", 11),
        fg="#e8dcc4",
        bg="#2d5016"
    )
    subtitle_label.pack()
    
    # Main content frame
    content_frame = tk.Frame(root, bg="#f5f1e8")
    content_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
    
    # Current folder section
    folder_section = tk.Frame(content_frame, bg="#ffffff", relief=tk.RIDGE, borderwidth=2)
    folder_section.pack(fill=tk.X, pady=(0, 15))
    
    config = load_config()
    current_folder = config.get("data_path", "Not set")
    
    tk.Label(
        folder_section, 
        text="📁 Current Data Folder:",
        font=("Arial", 10, "bold"),
        bg="#ffffff",
        fg="#2d5016"
    ).pack(anchor=tk.W, padx=10, pady=(10, 5))
    
    folder_label = tk.Label(
        folder_section, 
        text=current_folder,
        font=("Arial", 9),
        fg="#555555",
        bg="#ffffff",
        wraplength=580,
        justify=tk.LEFT
    )
    folder_label.pack(anchor=tk.W, padx=10, pady=(0, 10))
    
    def change_data_folder():
        """Allow user to change the data folder location"""
        config = load_config()
        messagebox.showinfo("Change Data Folder", 
            "Please select the new location of the PRINT_LABELS_BOXES_TINS folder.")
        
        path = filedialog.askdirectory(title="Select PRINT_LABELS_BOXES_TINS folder")
        
        if path:
            config["data_path"] = path
            save_config(config)
            messagebox.showinfo("Success", 
                "Data folder updated successfully!\n\n"
                "Please restart the application for changes to take effect.")
            root.destroy()
    
    change_folder_button = tk.Button(
        folder_section,
        text="Change Data Folder",
        command=change_data_folder,
        bg="#8b7355",  # Warm brown
        fg="white",
        font=("Arial", 9, "bold"),
        relief=tk.FLAT,
        padx=15,
        pady=5
    )
    change_folder_button.pack(pady=(0, 10))
    
    # Tea search section
    search_frame = tk.Frame(content_frame, bg="#f5f1e8")
    search_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
    
    tk.Label(
        search_frame,
        text="🍵 Search Tea Product:",
        font=("Arial", 11, "bold"),
        bg="#f5f1e8",
        fg="#2d5016"
    ).pack(anchor=tk.W, pady=(0, 5))
    
    search_var = tk.StringVar()
    search_entry = tk.Entry(
        search_frame,
        textvariable=search_var,
        width=50,
        font=("Arial", 11),
        relief=tk.SOLID,
        borderwidth=1
    )
    search_entry.pack(pady=5, ipady=5)
    
    # Custom styled listbox frame
    listbox_frame = tk.Frame(search_frame, bg="#ffffff", relief=tk.SOLID, borderwidth=1)
    listbox_frame.pack(pady=5, fill=tk.BOTH, expand=True)
    
    scrollbar = tk.Scrollbar(listbox_frame)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    search_listbox = tk.Listbox(
        listbox_frame,
        width=50,
        height=8,
        font=("Arial", 10),
        yscrollcommand=scrollbar.set,
        relief=tk.FLAT,
        selectbackground="#2d5016",
        selectforeground="white"
    )
    search_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scrollbar.config(command=search_listbox.yview)
    
    def update_search(event):
        """Update search results as user types"""
        search_term = search_var.get().lower()
        search_listbox.delete(0, tk.END)
        if search_term:
            matches = [tea for tea in products if search_term in tea.lower()]
            for tea in matches[:50]:  # Limit to 50 results
                search_listbox.insert(tk.END, tea)
    
    def on_select(event):
        """Handle tea selection from list"""
        selected_index = search_listbox.curselection()
        if selected_index:
            search_var.set(search_listbox.get(selected_index))
    
    search_entry.bind("<KeyRelease>", update_search)
    search_listbox.bind("<<ListboxSelect>>", on_select)
    
    # Bottom section - Quantity and Printer
    bottom_frame = tk.Frame(content_frame, bg="#f5f1e8")
    bottom_frame.pack(fill=tk.X, pady=(15, 0))
    
    # Quantity and Printer in horizontal layout
    controls_frame = tk.Frame(bottom_frame, bg="#f5f1e8")
    controls_frame.pack(pady=5)
    
    # Quantity
    qty_frame = tk.Frame(controls_frame, bg="#f5f1e8")
    qty_frame.pack(side=tk.LEFT, padx=20)
    
    tk.Label(
        qty_frame,
        text="Quantity:",
        font=("Arial", 11, "bold"),
        bg="#f5f1e8",
        fg="#2d5016"
    ).pack(side=tk.LEFT, padx=(0, 10))
    
    qty_var = tk.StringVar(value="1")
    qty_entry = tk.Entry(
        qty_frame,
        textvariable=qty_var,
        width=8,
        font=("Arial", 12),
        relief=tk.SOLID,
        borderwidth=1,
        justify=tk.CENTER
    )
    qty_entry.pack(side=tk.LEFT, ipady=5)
    
    # Printer
    printer_frame = tk.Frame(controls_frame, bg="#f5f1e8")
    printer_frame.pack(side=tk.LEFT, padx=20)
    
    tk.Label(
        printer_frame,
        text="Printer:",
        font=("Arial", 11, "bold"),
        bg="#f5f1e8",
        fg="#2d5016"
    ).pack(side=tk.LEFT, padx=(0, 10))
    
    try:
        printers = [printer[2] for printer in win32print.EnumPrinters(2)]
        if not printers:
            raise Exception("No printers found")
    except Exception as e:
        print(f"Error getting printers: {e}")
        printers = ["No printers available"]
    
    printer_combo = ttk.Combobox(
        printer_frame,
        values=printers,
        state="readonly",
        width=35,
        font=("Arial", 10)
    )
    if printers:
        printer_combo.current(0)
    printer_combo.pack(side=tk.LEFT)
    
    def on_print():
        """Handle print button click"""
        tea = search_var.get()
        if tea not in products:
            messagebox.showerror("Error", 
                "Please select a valid tea from the search results.")
            return
        
        try:
            quantity = int(qty_var.get())
            if quantity < 1:
                raise ValueError("Quantity must be at least 1")
            if quantity > 100:
                confirm = messagebox.askyesno("Confirm", 
                    f"You are about to print {quantity} labels. Continue?")
                if not confirm:
                    return
        except ValueError:
            messagebox.showerror("Error", "Quantity must be a valid number (1-100).")
            return
        
        printer_name = printer_combo.get()
        if not printer_name or printer_name == "No printers available":
            messagebox.showerror("Error", "Please select a valid printer.")
            return
        
        generate_label(tea, quantity, printer_name, df, LABEL_DIR, BARCODE_DIR)
    
    # Print button - prominent and centered
    print_button = tk.Button(
        bottom_frame,
        text="PRINT",
        command=on_print,
        bg="#2d5016",  # Deep tea green
        fg="white",
        font=("Arial", 16, "bold"),
        relief=tk.FLAT,
        padx=60,
        pady=15,
        cursor="hand2"
    )
    print_button.pack(pady=20)
    
    # Hover effect for print button
    def on_enter(e):
        print_button['bg'] = '#3d6826'
    
    def on_leave(e):
        print_button['bg'] = '#2d5016'
    
    print_button.bind("<Enter>", on_enter)
    print_button.bind("<Leave>", on_leave)
    
    # Info label below print button
    info_frame = tk.Frame(bottom_frame, bg="#fffbf0", relief=tk.SOLID, borderwidth=1)
    info_frame.pack(pady=(0, 10))
    
    info_label = tk.Label(
        info_frame,
        text="ℹ️  Quantity Entered is Labels\nEach Sheet has 20 Labels",
        font=("Arial", 10),
        fg="#8b7355",
        bg="#fffbf0",
        justify=tk.CENTER
    )
    info_label.pack(padx=20, pady=10)
    
    root.mainloop()

except Exception as e:
    print("Application failed to start.")
    print(traceback.format_exc())
    messagebox.showerror("Startup Error", 
        f"Application could not start:\n\n{str(e)}\n\n"
        f"Please check that:\n"
        f"1. The data folder is accessible\n"
        f"2. The Excel file exists\n"
        f"3. You have proper permissions"
    )