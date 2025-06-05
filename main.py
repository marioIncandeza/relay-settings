import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from rdb import gen_settings


class SettingsGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("SEL Settings Generator")
        self.root.geometry("600x400")
        
        # Variables to store paths and selection
        self.xl_path = tk.StringVar()
        self.template_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.selected_type = None
        self.workbook_params = None
        self.workbook_data = {
            'feeder': {'sheet_name': 'Feeder_351S', 'class_table': 'class_351S', 'settings_table': 'settings_351S'},
            'hv': {'sheet_name': 'HV_351S', 'class_table': 'class_HV351S', 'settings_table': 'settings_HV351S'},
            'xfmr_487E': {'sheet_name': 'XFMR_487E', 'class_table': 'class_487E', 'settings_table': 'settings_487E'},
            'cap_487V': {'sheet_name': 'CAP_487V', 'class_table': 'class_487V', 'settings_table': 'settings_487V'}
        }
        
        self.show_selection_screen()
        
    def show_selection_screen(self):
        # Clear any existing widgets
        for widget in self.root.winfo_children():
            widget.destroy()
            
        # Create selection frame
        selection_frame = ttk.Frame(self.root, padding="20")
        selection_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        
        # Title
        title_label = ttk.Label(selection_frame, text="Select Relay Type", font=('Helvetica', 14, 'bold'))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        # Buttons
        style = ttk.Style()
        style.configure('Large.TButton', padding=10)
        
        feeder_btn = ttk.Button(selection_frame, text="Feeder 351S", style='Large.TButton',
                               command=lambda: self.on_type_selected('feeder'))
        feeder_btn.grid(row=1, column=0, padx=10, pady=10, sticky=(tk.W, tk.E))
        
        hv_btn = ttk.Button(selection_frame, text="HV 351S", style='Large.TButton',
                           command=lambda: self.on_type_selected('hv'))
        hv_btn.grid(row=1, column=1, padx=10, pady=10, sticky=(tk.W, tk.E))

        xfmr_btn = ttk.Button(selection_frame, text="XFMR 487E", style='Large.TButton',
                              command=lambda: self.on_type_selected('xfmr_487E'))
        xfmr_btn.grid(row=2, column=0, padx=10, pady=10, sticky=(tk.W, tk.E))

        cap_btn = ttk.Button(selection_frame, text="CAP 487V", style='Large.TButton',
                              command=lambda: self.on_type_selected('cap_487V'))
        cap_btn.grid(row=2, column=1, padx=10, pady=10, sticky=(tk.W, tk.E))
        
        # Configure grid weights
        selection_frame.grid_columnconfigure(0, weight=1)
        selection_frame.grid_columnconfigure(1, weight=1)

        # Reset template path
        self.template_path = tk.StringVar()
        
    def on_type_selected(self, relay_type):
        self.selected_type = relay_type
        self.workbook_params = self.workbook_data[relay_type]
        self.show_main_interface()
        
    def show_main_interface(self):
        # Clear any existing widgets
        for widget in self.root.winfo_children():
            widget.destroy()
            
        # Create main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Title showing selected type
        title_map = {
            'feeder': "Feeder 351S Settings",
            'hv': "HV 351S Settings",
            'xfmr_487E': "Transformer 487E Settings",
            'cap_487V': "Capacitor Bank 487V Settings"
        }
        title_text = title_map.get(self.selected_type, "Relay Settings")
        title_label = ttk.Label(main_frame, text=title_text, font=('Helvetica', 12, 'bold'))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 15))
        
        # Back button
        back_btn = ttk.Button(main_frame, text="← Back", command=self.show_selection_screen)
        back_btn.grid(row=0, column=2, sticky=tk.E)
        
        # Excel file selection
        ttk.Label(main_frame, text="Settings Workbook:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.xl_path, width=50).grid(row=1, column=1, padx=5)
        ttk.Button(main_frame, text="Browse...", command=self.browse_excel).grid(row=1, column=2)
        
        # Template directory selection
        ttk.Label(main_frame, text="RDB Template Directory:").grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.template_path, width=50).grid(row=2, column=1, padx=5)
        ttk.Button(main_frame, text="Browse...", command=self.browse_template).grid(row=2, column=2)
        
        # Output directory selection
        ttk.Label(main_frame, text="Output Directory:").grid(row=3, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.output_path, width=50).grid(row=3, column=1, padx=5)
        ttk.Button(main_frame, text="Browse...", command=self.browse_output).grid(row=3, column=2)
        
        # Generate button
        ttk.Button(main_frame, text="Generate Settings", command=self.generate_settings).grid(row=4, column=0, columnspan=3, pady=20)
        
        # Status label
        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        ttk.Label(main_frame, textvariable=self.status_var).grid(row=5, column=0, columnspan=3)
        
        # Configure grid weights
        main_frame.grid_columnconfigure(1, weight=1)

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
            
            # Choose the appropriate generation function
            gen_function = gen_settings  # _351S if self.selected_type == 'feeder' else gen_settings_HV351S
            
            gen_function(
                xl_path=self.xl_path.get(),
                template_path=self.template_path.get(),
                output_path=self.output_path.get(),
                workbook_params=self.workbook_params
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
