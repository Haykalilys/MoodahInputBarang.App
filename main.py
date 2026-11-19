import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
from datetime import datetime

class ProductPage(tk.Frame):
    def __init__(self, master):
        super().__init__(master, bg="#f5f7fa")
        self.search_var = tk.StringVar()
        self.sort_var = tk.StringVar(value="No")
        self.sort_order = tk.BooleanVar(value=True)
        
        self.setup_style()
        self.header()
        self.filter_section()
        self.table_section()
        self.load_sample_data()

    def setup_style(self):
        self.style = ttk.Style()
        self.style.theme_use("clam")
        
        # Color scheme
        self.bg_color = "#f5f7fa"
        self.primary_color = "#1a759f"
        self.secondary_color = "#1cc88a"
        self.danger_color = "#e74a3b"
        self.light_color = "#ffffff"
        self.text_color = "#5a5c69"
        
        # Configure styles
        self.style.configure("TFrame", background=self.bg_color)
        self.style.configure("TLabel", background=self.bg_color, foreground=self.text_color)
        
        # Treeview style
        self.style.configure("Treeview",
                            background=self.light_color,
                            foreground=self.text_color,
                            rowheight=30,
                            fieldbackground=self.light_color,
                            font=("Segoe UI", 10),
                            borderwidth=0,
                            relief="flat")
        self.style.map("Treeview", 
                      background=[("selected", self.primary_color)],
                      foreground=[("selected", "white")])
        
        # Treeview heading style
        self.style.configure("Treeview.Heading",
                            background=self.primary_color,
                            foreground="white",
                            font=("Segoe UI", 10, "bold"),
                            borderwidth=0,
                            relief="flat")
        
        # Button styles
        self.style.configure("Primary.TButton",
                            foreground="white",
                            background=self.primary_color,
                            font=("Segoe UI", 9, "bold"),
                            borderwidth=0,
                            focuscolor=self.bg_color,
                            focusthickness=0,
                            padding=6,
                            relief="flat")
        self.style.map("Primary.TButton",
                     background=[("active", "#168aad")])
        
        self.style.configure("Success.TButton",
                            foreground="white",
                            background=self.secondary_color,
                            font=("Segoe UI", 9, "bold"),
                            borderwidth=0,
                            padding=6,
                            relief="flat")
        self.style.map("Success.TButton",
                     background=[("active", "#17a673")])
        
        self.style.configure("Danger.TButton",
                            foreground="white",
                            background=self.danger_color,
                            font=("Segoe UI", 9, "bold"),
                            borderwidth=0,
                            padding=6,
                            relief="flat")
        self.style.map("Danger.TButton",
                     background=[("active", "#c23321")])
        
        # Combobox style
        self.style.configure("TCombobox",
                            fieldbackground=self.light_color,
                            foreground=self.text_color,
                            selectbackground=self.primary_color,
                            selectforeground="white",
                            insertcolor=self.text_color,
                            padding=5,
                            relief="flat")
        self.style.map("TCombobox",
                      fieldbackground=[("readonly", self.light_color)],
                      selectbackground=[("readonly", self.primary_color)])

    def header(self):
        header = tk.Frame(self, bg=self.light_color, height=80, padx=20)
        header.pack(fill=tk.X)
        
        # Title
        title_frame = tk.Frame(header, bg=self.light_color)
        title_frame.pack(side=tk.LEFT)
        
        tk.Label(title_frame, text="Manajemen Produk", bg=self.light_color, 
                font=("Segoe UI", 18, "bold"), fg=self.text_color).pack(side=tk.LEFT)
        
        # Action buttons
        action_frame = tk.Frame(header, bg=self.light_color)
        action_frame.pack(side=tk.RIGHT)
        

    def filter_section(self):
        filter_frame = tk.Frame(self, bg=self.bg_color, padx=20)
        filter_frame.pack(fill=tk.X, pady=10)
        
        # Search frame
        search_frame = tk.Frame(filter_frame, bg=self.bg_color)
        search_frame.pack(side=tk.LEFT, padx=(0, 10))
        
        search_entry = ttk.Entry(search_frame, textvariable=self.search_var, width=30)
        search_entry.pack(side=tk.LEFT, ipady=4)
        
        ttk.Button(search_frame, text="Cari", command=self.search_product,
                  style="Primary.TButton").pack(side=tk.LEFT, padx=5)
        
        # Action buttons frame
        action_frame = tk.Frame(filter_frame, bg=self.bg_color)
        action_frame.pack(side=tk.RIGHT)
        
        ttk.Button(action_frame, text="Export ke Excel", command=self.save_to_excel,
                  style="Success.TButton").pack(side=tk.LEFT, padx=5)
        
        # Add product button
        ttk.Button(action_frame, text="+ Tambah Produk", command=self.add_product_popup,
                  style="Primary.TButton").pack(side=tk.LEFT, padx=5)

    def table_section(self):
        container = tk.Frame(self, bg=self.bg_color, padx=20)
        container.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Treeview frame with shadow effect
        tree_container = tk.Frame(container, bg=self.bg_color)
        tree_container.pack(fill=tk.BOTH, expand=True)
        
        # Create a Treeview with scrollbars (but not showing them)
        tree_frame = tk.Frame(tree_container, bg=self.bg_color)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create Treeview
        self.tree = ttk.Treeview(tree_frame, 
                               columns=("No", "Nama", "ID", "Kategori", "Harga", "Stok", "Tanggal"), 
                               show="headings",
                               selectmode="extended")
        self.tree.pack(fill=tk.BOTH, expand=True)
        
        # Configure columns
        columns = {
            "No": {"width": 60, "anchor": tk.CENTER, "stretch": False},
            "Nama": {"width": 200, "anchor": tk.W, "stretch": True},
            "ID": {"width": 120, "anchor": tk.CENTER, "stretch": False},
            "Kategori": {"width": 120, "anchor": tk.CENTER, "stretch": False},
            "Harga": {"width": 100, "anchor": tk.E, "stretch": False},
            "Stok": {"width": 80, "anchor": tk.CENTER, "stretch": False},
            "Tanggal": {"width": 120, "anchor": tk.CENTER, "stretch": False}
        }
        
        for col, settings in columns.items():
            self.tree.heading(col, text=col, command=lambda c=col: self.sort_by_column(c))
            self.tree.column(col, width=settings["width"], anchor=settings["anchor"], 
                           stretch=settings["stretch"])
        
        # Enable scrolling without visible scrollbars
        self.tree.bind("<MouseWheel>", lambda e: self.tree.yview_scroll(-1*(e.delta//120), "units"))
        self.tree.bind("<Shift-MouseWheel>", lambda e: self.tree.xview_scroll(-1*(e.delta//120), "units"))
        
        # Button frame
        button_frame = tk.Frame(container, bg=self.bg_color)
        button_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(button_frame, text="Hapus Pilihan", command=self.delete_selected,
                  style="Danger.TButton").pack(side=tk.RIGHT)

    def sort_by_column(self, column):
        current_sort = self.sort_var.get()
        if current_sort == column:
            self.sort_order.set(not self.sort_order.get())
        else:
            self.sort_var.set(column)
            self.sort_order.set(True)
        
        self.sort_products()

    def sort_products(self):
        sort_column = self.sort_var.get()
        reverse_order = not self.sort_order.get()
        
        # Get all items and their values
        items = [(self.tree.set(item, sort_column), item) for item in self.tree.get_children('')]
        
        # Try to convert to numbers for proper numeric sorting
        try:
            items.sort(key=lambda x: float(x[0].replace(',', '')) if x[0] else 0, reverse=reverse_order)
        except ValueError:
            items.sort(key=lambda x: x[0], reverse=reverse_order)
        
        # Reorder items in treeview
        for index, (val, item) in enumerate(items):
            self.tree.move(item, '', index)
        
        # Apply alternating row colors after sorting
        self.apply_alternating_row_colors()

    def load_sample_data(self):
        # Clear existing data
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Add sample data with categories
        sample_data = [
            [1, "Laptop ASUS X441", "PRD001", "Elektronik", "8,500,000", 15, "2023-05-10"],
            [2, "Smartphone Samsung A53", "PRD002", "Elektronik", "4,200,000", 32, "2023-05-12"],
            [3, "Monitor LG 24MK430", "PRD003", "Elektronik", "2,300,000", 8, "2023-05-15"],
            [4, "Keyboard Mechanical", "PRD004", "Aksesoris", "750,000", 45, "2023-05-18"],
            [5, "Mouse Wireless Logitech", "PRD005", "Aksesoris", "350,000", 60, "2023-05-20"],
            [6, "Headphones Sony WH-1000XM4", "PRD006", "Elektronik", "3,500,000", 12, "2023-05-22"],
            [7, "External SSD 1TB", "PRD007", "Penyimpanan", "1,200,000", 20, "2023-05-25"],
            [8, "Webcam HD Pro", "PRD008", "Aksesoris", "650,000", 18, "2023-05-28"],
            [9, "Bluetooth Speaker", "PRD009", "Elektronik", "950,000", 25, "2023-05-30"],
            [10, "USB-C Hub", "PRD010", "Aksesoris", "450,000", 30, "2023-06-01"]
        ]
        
        for product in sample_data:
            self.tree.insert("", "end", values=product)
        
        # Apply alternating row colors
        self.apply_alternating_row_colors()

    def apply_alternating_row_colors(self):
        for i, item in enumerate(self.tree.get_children()):
            if i % 2 == 0:
                self.tree.tag_configure('evenrow', background='#f5f7fa')
                self.tree.item(item, tags=('evenrow',))
            else:
                self.tree.tag_configure('oddrow', background=self.light_color)
                self.tree.item(item, tags=('oddrow',))

    def add_product_popup(self):
        win = tk.Toplevel(self)
        win.title("Tambah Produk Baru")
        win.geometry("450x500")
        win.configure(bg=self.bg_color)
        win.resizable(False, False)
        
        # Center the window
        window_width = 450
        window_height = 500
        screen_width = win.winfo_screenwidth()
        screen_height = win.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        win.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
        # Form container
        form_frame = tk.Frame(win, bg=self.bg_color, padx=20, pady=20)
        form_frame.pack(fill=tk.BOTH, expand=True)
        
        # Form title
        tk.Label(form_frame, text="Tambah Produk Baru", bg=self.bg_color,
                font=("Segoe UI", 12, "bold"), fg=self.text_color).pack(pady=(0, 10))
        
        fields = [
            {"label": "Nama Produk", "type": "entry"},
            {"label": "ID Produk", "type": "entry"},
            {"label": "Kategori", "type": "combobox", "values": ["Elektronik", "Aksesoris", "Penyimpanan", "Perangkat Jaringan"]},
            {"label": "Harga", "type": "entry"},
            {"label": "Stok", "type": "entry"}
        ]
        
        entries = {}
        
        for field in fields:
            field_frame = tk.Frame(form_frame, bg=self.bg_color)
            field_frame.pack(fill=tk.X, pady=8)
            
            tk.Label(field_frame, text=field["label"], bg=self.bg_color,
                    font=("Segoe UI", 9), fg=self.text_color).pack(anchor=tk.W)
            
            if field["type"] == "entry":
                entry = ttk.Entry(field_frame, font=("Segoe UI", 10))
                entry.pack(fill=tk.X, ipady=4)
            elif field["type"] == "combobox":
                entry = ttk.Combobox(field_frame, values=field["values"], font=("Segoe UI", 10), state="readonly")
                entry.pack(fill=tk.X, ipady=4)
            
            entries[field["label"]] = entry
        
        # Button frame
        button_frame = tk.Frame(form_frame, bg=self.bg_color, pady=20)
        button_frame.pack(fill=tk.X)
        
        def save():
            values = [
                len(self.tree.get_children()) + 1,  # Auto number
                entries["Nama Produk"].get(),
                entries["ID Produk"].get(),
                entries["Kategori"].get(),
                entries["Harga"].get(),
                entries["Stok"].get(),
                pd.Timestamp.now().strftime("%Y-%m-%d")  # Current date
            ]
            
            if not all(values[1:5]):  # Skip checking auto number and date
                messagebox.showwarning("Warning", "Harap isi semua field yang diperlukan.")
                return
                
            self.tree.insert("", "end", values=values)
            # Reapply alternating colors after adding new item
            self.apply_alternating_row_colors()
            messagebox.showinfo("Success", "Produk berhasil ditambahkan.")
            win.destroy()
        
        ttk.Button(button_frame, text="Simpan Produk", command=save,
                  style="Primary.TButton").pack(side=tk.RIGHT)
        ttk.Button(button_frame, text="Batal", command=win.destroy,
                  style="Danger.TButton").pack(side=tk.RIGHT, padx=10)

    def delete_selected(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Silakan pilih setidaknya satu produk untuk dihapus.")
            return
            
        confirm = messagebox.askyesno("Confirm Delete", 
                                    f"Apakah Anda yakin ingin menghapus {len(selected)} produk yang dipilih?")
        if confirm:
            for item in selected:
                self.tree.delete(item)
            # Reapply alternating colors after deletion
            self.apply_alternating_row_colors()
            messagebox.showinfo("Success", f"{len(selected)} produk berhasil dihapus.")

    def save_to_excel(self):
        data = []
        for child in self.tree.get_children():
            data.append(self.tree.item(child)["values"])
            
        if not data:
            messagebox.showwarning("Warning", "Tidak ada data untuk diekspor.")
            return
            
        df = pd.DataFrame(data, columns=["No", "Nama", "ID", "Kategori", "Harga", "Stok", "Tanggal"])
        try:
            df.to_excel("products_inventory.xlsx", index=False)
            messagebox.showinfo("Success", "Data berhasil diekspor ke products_inventory.xlsx")
        except Exception as e:
            messagebox.showerror("Error", f"Gagal mengekspor data: {str(e)}")

    def search_product(self):
        keyword = self.search_var.get().lower()
        if not keyword:
            messagebox.showwarning("Warning", "Silakan masukkan kata kunci pencarian.")
            return
            
        found_items = []
        for item in self.tree.get_children():
            values = self.tree.item(item)["values"]
            if any(keyword in str(val).lower() for val in values):
                found_items.append(item)
                
        if found_items:
            self.tree.selection_set(found_items)
            self.tree.see(found_items[0])
        else:
            messagebox.showinfo("Not Found", "Tidak ditemukan produk yang sesuai.")

# Main application
if __name__ == "__main__":
    root = tk.Tk()
    root.title("Moodah Input Barang")
    root.geometry("1000x650")
    root.configure(bg="#f5f7fa")
    
    # Center the main window
    window_width = 1000
    window_height = 650
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width - window_width) // 2
    y = (screen_height - window_height) // 2
    root.geometry(f"{window_width}x{window_height}+{x}+{y}")
    
    app = ProductPage(root)
    app.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    
    root.mainloop()