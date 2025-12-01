import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import mysql.connector
import pickle
import os
import glob
from datetime import datetime

class JewelryBillingAutomation:
    def __init__(self, root):
        self.root = root
        self.root.title("Jewelry Shop Billing Automation System")
        self.root.geometry("900x700")
        self.root.configure(bg='#f0f0f0')
        
        # Configuration storage
        self.config = {
            'folder_path': '',
            'db_host': 'localhost',
            'db_user': 'root',
            'db_password': '',
            'db_name': 'jewelry_shop'
        }
        
        self.load_config()
        self.create_ui()
        
    def load_config(self):
        """Load saved configuration from pickle file"""
        if os.path.exists('config.pkl'):
            with open('config.pkl', 'rb') as f:
                self.config = pickle.load(f)
                
    def save_config(self):
        """Save configuration to pickle file"""
        with open('config.pkl', 'wb') as f:
            pickle.dump(self.config, f)
    
    def create_ui(self):
        """Create the user interface"""
        # Header
        header_frame = tk.Frame(self.root, bg='#2c3e50', height=80)
        header_frame.pack(fill='x')
        header_frame.pack_propagate(False)
        
        title_label = tk.Label(header_frame, text="💎 Jewelry Shop Billing Automation",
                               font=('Arial', 24, 'bold'), bg='#2c3e50', fg='white')
        title_label.pack(pady=20)
        
        # Main container
        main_frame = tk.Frame(self.root, bg='#f0f0f0')
        main_frame.pack(fill='both', expand=True, padx=20, pady=20)
        
        # Step 1: Folder Selection
        step1_frame = tk.LabelFrame(main_frame, text="Step 1: Select Monthly Excel Files Folder",
                                   font=('Arial', 12, 'bold'), bg='white', fg='#2c3e50',
                                   padx=15, pady=15)
        step1_frame.pack(fill='x', pady=(0, 15))
        
        folder_frame = tk.Frame(step1_frame, bg='white')
        folder_frame.pack(fill='x')
        
        self.folder_entry = tk.Entry(folder_frame, font=('Arial', 10), width=50)
        self.folder_entry.pack(side='left', padx=(0, 10))
        self.folder_entry.insert(0, self.config.get('folder_path', ''))
        
        browse_btn = tk.Button(folder_frame, text="📁 Browse", command=self.browse_folder,
                              bg='#3498db', fg='white', font=('Arial', 10, 'bold'),
                              padx=15, pady=5, cursor='hand2')
        browse_btn.pack(side='left')
        
        # Step 2: Database Configuration
        step2_frame = tk.LabelFrame(main_frame, text="Step 2: MySQL Database Configuration",
                                   font=('Arial', 12, 'bold'), bg='white', fg='#2c3e50',
                                   padx=15, pady=15)
        step2_frame.pack(fill='x', pady=(0, 15))
        
        db_grid = tk.Frame(step2_frame, bg='white')
        db_grid.pack(fill='x')
        
        # Database fields
        tk.Label(db_grid, text="Host:", bg='white', font=('Arial', 10)).grid(row=0, column=0, sticky='w', pady=5)
        self.host_entry = tk.Entry(db_grid, font=('Arial', 10), width=20)
        self.host_entry.grid(row=0, column=1, padx=10, pady=5, sticky='w')
        self.host_entry.insert(0, self.config.get('db_host', 'localhost'))
        
        tk.Label(db_grid, text="User:", bg='white', font=('Arial', 10)).grid(row=0, column=2, sticky='w', pady=5)
        self.user_entry = tk.Entry(db_grid, font=('Arial', 10), width=20)
        self.user_entry.grid(row=0, column=3, padx=10, pady=5, sticky='w')
        self.user_entry.insert(0, self.config.get('db_user', 'root'))
        
        tk.Label(db_grid, text="Password:", bg='white', font=('Arial', 10)).grid(row=1, column=0, sticky='w', pady=5)
        self.pass_entry = tk.Entry(db_grid, font=('Arial', 10), width=20, show='*')
        self.pass_entry.grid(row=1, column=1, padx=10, pady=5, sticky='w')
        self.pass_entry.insert(0, self.config.get('db_password', ''))
        
        tk.Label(db_grid, text="Database:", bg='white', font=('Arial', 10)).grid(row=1, column=2, sticky='w', pady=5)
        self.db_entry = tk.Entry(db_grid, font=('Arial', 10), width=20)
        self.db_entry.grid(row=1, column=3, padx=10, pady=5, sticky='w')
        self.db_entry.insert(0, self.config.get('db_name', 'jewelry_shop'))
        
        test_btn = tk.Button(step2_frame, text="🔌 Test Connection", command=self.test_connection,
                           bg='#9b59b6', fg='white', font=('Arial', 10, 'bold'),
                           padx=15, pady=5, cursor='hand2')
        test_btn.pack(pady=(10, 0))
        
        # Step 3: Preview
        step3_frame = tk.LabelFrame(main_frame, text="Step 3: Data Preview",
                                   font=('Arial', 12, 'bold'), bg='white', fg='#2c3e50',
                                   padx=15, pady=15)
        step3_frame.pack(fill='both', expand=True, pady=(0, 15))
        
        preview_btn = tk.Button(step3_frame, text="👁️ Preview Data", command=self.preview_data,
                              bg='#e67e22', fg='white', font=('Arial', 10, 'bold'),
                              padx=15, pady=5, cursor='hand2')
        preview_btn.pack(pady=(0, 10))
        
        # Treeview for data preview
        tree_frame = tk.Frame(step3_frame, bg='white')
        tree_frame.pack(fill='both', expand=True)
        
        tree_scroll = tk.Scrollbar(tree_frame)
        tree_scroll.pack(side='right', fill='y')
        
        self.tree = ttk.Treeview(tree_frame, yscrollcommand=tree_scroll.set, height=8)
        self.tree.pack(fill='both', expand=True)
        tree_scroll.config(command=self.tree.yview)
        
        # Info label
        self.info_label = tk.Label(step3_frame, text="", bg='white', 
                                   font=('Arial', 9), fg='#7f8c8d')
        self.info_label.pack(pady=(5, 0))
        
        # Step 4: Process
        step4_frame = tk.Frame(main_frame, bg='#f0f0f0')
        step4_frame.pack(fill='x')
        
        process_btn = tk.Button(step4_frame, text="⚙️ Process & Save to Database",
                               command=self.process_data, bg='#27ae60', fg='white',
                               font=('Arial', 12, 'bold'), padx=30, pady=10, cursor='hand2')
        process_btn.pack(side='left', padx=(0, 10))
        
        export_btn = tk.Button(step4_frame, text="📊 Export Yearly Excel",
                              command=self.export_yearly_excel, bg='#16a085', fg='white',
                              font=('Arial', 12, 'bold'), padx=30, pady=10, cursor='hand2')
        export_btn.pack(side='left')
        
        # Status bar
        self.status_label = tk.Label(self.root, text="Ready", bg='#34495e', fg='white',
                                     font=('Arial', 9), anchor='w', padx=10)
        self.status_label.pack(side='bottom', fill='x')
    
    def browse_folder(self):
        """Browse for folder containing Excel files"""
        folder = filedialog.askdirectory()
        if folder:
            self.folder_entry.delete(0, tk.END)
            self.folder_entry.insert(0, folder)
            self.config['folder_path'] = folder
            self.save_config()
            self.status_label.config(text=f"Folder selected: {folder}")
    
    def test_connection(self):
        """Test MySQL database connection"""
        try:
            conn = mysql.connector.connect(
                host=self.host_entry.get(),
                user=self.user_entry.get(),
                password=self.pass_entry.get()
            )
            
            # Create database if not exists
            cursor = conn.cursor()
            cursor.execute(f"CREATE DATABASE IF NOT EXISTS {self.db_entry.get()}")
            conn.commit()
            cursor.close()
            conn.close()
            
            # Update config
            self.config['db_host'] = self.host_entry.get()
            self.config['db_user'] = self.user_entry.get()
            self.config['db_password'] = self.pass_entry.get()
            self.config['db_name'] = self.db_entry.get()
            self.save_config()
            
            messagebox.showinfo("Success", "Database connection successful!\nDatabase created if not existed.")
            self.status_label.config(text="Database connection successful")
        except Exception as e:
            messagebox.showerror("Error", f"Connection failed:\n{str(e)}")
            self.status_label.config(text="Database connection failed")
    
    def preview_data(self):
        """Preview data from Excel files"""
        folder = self.folder_entry.get()
        if not folder or not os.path.exists(folder):
            messagebox.showwarning("Warning", "Please select a valid folder first!")
            return
        
        try:
            # Get all Excel files
            excel_files = glob.glob(os.path.join(folder, "*.xlsx"))
            if not excel_files:
                messagebox.showwarning("Warning", "No Excel files found in selected folder!")
                return
            
            # Read first file for preview
            df = pd.read_excel(excel_files[0])
            
            # Clear existing tree
            self.tree.delete(*self.tree.get_children())
            
            # Configure columns
            self.tree['columns'] = list(df.columns)
            self.tree['show'] = 'headings'
            
            for col in df.columns:
                self.tree.heading(col, text=col)
                self.tree.column(col, width=100)
            
            # Add data (first 10 rows)
            for idx, row in df.head(10).iterrows():
                self.tree.insert('', 'end', values=list(row))
            
            self.info_label.config(text=f"Found {len(excel_files)} Excel files | Showing first 10 rows from {os.path.basename(excel_files[0])}")
            self.status_label.config(text=f"Preview loaded: {len(excel_files)} files found")
        except Exception as e:
            messagebox.showerror("Error", f"Preview failed:\n{str(e)}")
    
    def process_data(self):
        """Process all Excel files and save to database"""
        folder = self.folder_entry.get()
        if not folder or not os.path.exists(folder):
            messagebox.showwarning("Warning", "Please select a valid folder first!")
            return
        
        try:
            self.status_label.config(text="Processing data...")
            self.root.update()
            
            # Get all Excel files
            excel_files = glob.glob(os.path.join(folder, "*.xlsx"))
            
            # Read and combine all files
            all_data = []
            for file in excel_files:
                df = pd.read_excel(file)
                all_data.append(df)
            
            combined_df = pd.concat(all_data, ignore_index=True)
            
            # Data cleaning
            combined_df = combined_df.drop_duplicates()
            combined_df = combined_df.dropna(subset=['Bill_No'])
            
            # Connect to database
            conn = mysql.connector.connect(
                host=self.config['db_host'],
                user=self.config['db_user'],
                password=self.config['db_password'],
                database=self.config['db_name']
            )
            cursor = conn.cursor()
            
            # Create table
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS billing_records (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    bill_no VARCHAR(50) UNIQUE,
                    date DATE,
                    customer_name VARCHAR(100),
                    contact_number VARCHAR(20),
                    item_name VARCHAR(100),
                    quantity INT,
                    weight_grams DECIMAL(10, 2),
                    rate_per_gram DECIMAL(10, 2),
                    making_charges DECIMAL(10, 2),
                    total_amount DECIMAL(12, 2),
                    payment_mode VARCHAR(50),
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)
            
            # Insert data
            insert_count = 0
            for _, row in combined_df.iterrows():
                try:
                    cursor.execute("""
                        INSERT INTO billing_records 
                        (bill_no, date, customer_name, contact_number, item_name, 
                         quantity, weight_grams, rate_per_gram, making_charges, 
                         total_amount, payment_mode)
                        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                        ON DUPLICATE KEY UPDATE
                        date = VALUES(date),
                        customer_name = VALUES(customer_name)
                    """, (
                        row['Bill_No'], row['Date'], row['Customer_Name'],
                        row['Contact_Number'], row['Item_Name'], row['Quantity'],
                        row['Weight_Grams'], row['Rate_Per_Gram'], row['Making_Charges'],
                        row['Total_Amount'], row['Payment_Mode']
                    ))
                    insert_count += 1
                except:
                    continue
            
            conn.commit()
            cursor.close()
            conn.close()
            
            messagebox.showinfo("Success", 
                f"Data processed successfully!\n\n"
                f"Files processed: {len(excel_files)}\n"
                f"Records inserted/updated: {insert_count}\n"
                f"Duplicates removed: {len(combined_df) - insert_count}")
            self.status_label.config(text=f"Processing complete: {insert_count} records saved")
        except Exception as e:
            messagebox.showerror("Error", f"Processing failed:\n{str(e)}")
            self.status_label.config(text="Processing failed")
    
    def export_yearly_excel(self):
        """Export consolidated yearly Excel file"""
        try:
            self.status_label.config(text="Exporting yearly Excel...")
            self.root.update()
            
            # Connect to database
            conn = mysql.connector.connect(
                host=self.config['db_host'],
                user=self.config['db_user'],
                password=self.config['db_password'],
                database=self.config['db_name']
            )
            
            # Read all data
            query = "SELECT * FROM billing_records ORDER BY date"
            df = pd.read_sql(query, conn)
            conn.close()
            
            if df.empty:
                messagebox.showwarning("Warning", "No data found in database!")
                return
            
            # Save file
            year = datetime.now().year
            filename = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                initialfile=f"Jewelry_Billing_Yearly_{year}.xlsx",
                filetypes=[("Excel files", "*.xlsx")]
            )
            
            if filename:
                # Create Excel writer with multiple sheets
                with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                    # All data
                    df.to_excel(writer, sheet_name='All Records', index=False)
                    
                    # Summary by month
                    df['date'] = pd.to_datetime(df['date'])
                    monthly_summary = df.groupby(df['date'].dt.month).agg({
                        'total_amount': 'sum',
                        'bill_no': 'count'
                    }).rename(columns={'bill_no': 'total_transactions'})
                    monthly_summary.to_excel(writer, sheet_name='Monthly Summary')
                    
                    # Top customers
                    top_customers = df.groupby('customer_name')['total_amount'].sum().sort_values(ascending=False).head(10)
                    top_customers.to_excel(writer, sheet_name='Top Customers')
                
                messagebox.showinfo("Success", f"Yearly Excel file exported successfully!\n\nLocation: {filename}")
                self.status_label.config(text=f"Export complete: {filename}")
        except Exception as e:
            messagebox.showerror("Error", f"Export failed:\n{str(e)}")
            self.status_label.config(text="Export failed")

if __name__ == "__main__":
    root = tk.Tk()
    app = JewelryBillingAutomation(root)
    root.mainloop()