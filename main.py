import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from rdb import gen_settings_351S

class SettingsGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("SEL Settings Generator")
        self.root.geometry("600x300")
        
        # Variables to store paths
        self.xl_path = tk.StringVar()
        self.template_path = tk.StringVar()
        self.output_path = tk.StringVar()
        
        # Create main frame
        main_frame = ttk.Frame(root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Excel file selection
        ttk.Label(main_frame, text="Settings Workbook:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.xl_path, width=50).grid(row=0, column=1, padx=5)
        ttk.Button(main_frame, text="Browse...", command=self.browse_excel).grid(row=0, column=2)
        
        # Template directory selection
        ttk.Label(main_frame, text="RDB Template Directory:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.template_path, width=50).grid(row=1, column=1, padx=5)
        ttk.Button(main_frame, text="Browse...", command=self.browse_template).grid(row=1, column=2)
        
        # Output directory selection
        ttk.Label(main_frame, text="Output Directory:").grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.output_path, width=50).grid(row=2, column=1, padx=5)
        ttk.Button(main_frame, text="Browse...", command=self.browse_output).grid(row=2, column=2)
        
        # Generate button
        ttk.Button(main_frame, text="Generate Settings", command=self.generate_settings).grid(row=3, column=0, columnspan=3, pady=20)
        
        # Status label
        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        ttk.Label(main_frame, textvariable=self.status_var).grid(row=4, column=0, columnspan=3)

    def browse_excel(self):
        filename = filedialog.askopenfilename(
            title="Select Settings Workbook",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.xl_path.set(filename)

    def browse_template(self):
        dirname = filedialog.askdirectory(title="Select RDB Template Directory")
        if dirname:
            self.template_path.set(dirname)

    def browse_output(self):
        dirname = filedialog.askdirectory(title="Select Output Directory")
        if dirname:
            self.output_path.set(dirname)

    def generate_settings(self):
        # Validate inputs
        if not all([self.xl_path.get(), self.template_path.get(), self.output_path.get()]):
            messagebox.showerror("Error", "Please select all required paths")
            return

        try:
            self.status_var.set("Generating settings...")
            self.root.update()
            
            gen_settings_351S(
                xl_path=self.xl_path.get(),
                template_path=self.template_path.get(),
                output_path=self.output_path.get()
            )
            
            self.status_var.set("Settings generated successfully!")
            messagebox.showinfo("Success", "Settings have been generated successfully!")
            
        except Exception as e:
            self.status_var.set("Error occurred during generation")
            messagebox.showerror("Error", f"An error occurred:\n{str(e)}")

def main():
    root = tk.Tk()
    app = SettingsGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
