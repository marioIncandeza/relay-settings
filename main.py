import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from rdb import gen_settings


class SettingsGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("SEL Settings Generator")
        self.root.geometry("650x500")

        # --- Persistent Variables (Workbook & Output) ---
        # These remain set even if you go Back and select a different relay
        self.xl_path = tk.StringVar()
        self.output_path = tk.StringVar()

        # --- Non-Persistent Variable (Template) ---
        # We initialize it here, but we will force-clear it later
        self.template_path = tk.StringVar()

        self.include_comments = tk.BooleanVar(value=True)
        self.status_var = tk.StringVar(value="Ready")
        self.selected_type = None
        self.workbook_params = None
        self.region_vars = {}

        # Relay Configuration Data
        self.relay_config = {
            'feeder': {'label': 'Feeder 351S', 'params': {'sheet_name': 'FDR_351S', 'class_table': 'class_351S',
                                                          'settings_table': 'settings_351S'}},
            'hv': {'label': 'HV 351S', 'params': {'sheet_name': 'HV_351S', 'class_table': 'class_HV351S',
                                                  'settings_table': 'settings_HV351S'}},
            'xfmr_487E': {'label': 'XFMR 487E', 'params': {'sheet_name': 'XFMR_487E', 'class_table': 'class_487E',
                                                           'settings_table': 'settings_487E'}},
            'cap_487V': {'label': 'CAP 487V', 'params': {'sheet_name': 'CAP_487V', 'class_table': 'class_487V',
                                                         'settings_table': 'settings_487V'}},
            'bus_587Z': {'label': 'BUS 587Z', 'params': {'sheet_name': 'BUS_587Z', 'class_table': 'class_587Z',
                                                         'settings_table': 'settings_587Z'}},
            'mtr_735': {'label': 'MTR 735', 'params': {'sheet_name': 'MTR_735', 'class_table': 'class_735',
                                                       'settings_table': 'settings_735'}},
            'dpac_2440': {'label': 'DPAC 2440', 'params': {'sheet_name': 'DPAC_2440', 'class_table': 'class_2440',
                                                           'settings_table': 'settings_2440'}},
            'xfmr_787': {'label': 'XFMR 787', 'params': {'sheet_name': 'XFMR_787', 'class_table': 'class_787',
                                                         'settings_table': 'settings_787'}},
            'line_411L': {'label': 'LINE 411L', 'params': {'sheet_name': 'Line_411L', 'class_table': 'class_411L',
                                                           'settings_table': 'settings_411L'}}
        }

        # Shared definition for 351S style groups
        common_351_groups = {
            "labels": [f"Group {i}" for i in range(1, 7)] + [f"Logic {i}" for i in range(1, 7)],
            "shorthand": {**{f"Group {i}": str(i) for i in range(1, 7)}, **{f"Logic {i}": f"L{i}" for i in range(1, 7)}}
        }

        common_400_groups = {
            "labels": [f"Set {i}" for i in range(1, 7)] + [f"Protection Logic {i}" for i in range(1, 7)],
            "shorthand": {**{f"Set {i}": f"S{i}" for i in range(1, 7)},
                          **{f"Protection Logic {i}": f"L{i}" for i in range(1, 7)}}
        }

        common_787_groups = {
            "labels": [f"Set {i}" for i in range(1, 5)] + [f"Logic {i}" for i in range(1, 5)],
            "shorthand": {**{f"Set {i}": str(i) for i in range(1, 5)}, **{f"Logic {i}": f"L{i}" for i in range(1, 5)}}
        }

        self.relay_region_metadata = {
            "feeder": common_351_groups,
            "hv": common_351_groups,
            "xfmr_487E": common_400_groups,
            "cap_487V": common_400_groups,
            "xfmr_787": common_787_groups,
            "line_411L": common_400_groups
        }

        self.show_selection_screen()

    def show_selection_screen(self):
        for widget in self.root.winfo_children():
            widget.destroy()

        selection_frame = ttk.Frame(self.root, padding="20")
        selection_frame.grid(row=0, column=0, sticky="nsew")
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

        ttk.Label(selection_frame, text="Select Relay Type", font=('Helvetica', 14, 'bold')).grid(row=0, column=0,
                                                                                                  columnspan=2,
                                                                                                  pady=(0, 20))

        style = ttk.Style()
        style.configure('Large.TButton', padding=10)

        keys = list(self.relay_config.keys())
        for index, key in enumerate(keys):
            row = (index // 2) + 1
            col = index % 2
            btn = ttk.Button(
                selection_frame,
                text=self.relay_config[key]['label'],
                style='Large.TButton',
                command=lambda k=key: self.on_type_selected(k)
            )
            btn.grid(row=row, column=col, padx=10, pady=10, sticky="ew")

        selection_frame.grid_columnconfigure(0, weight=1)
        selection_frame.grid_columnconfigure(1, weight=1)

    def on_type_selected(self, relay_type):
        self.selected_type = relay_type
        self.workbook_params = self.relay_config[relay_type]['params']

        # --- KEY CHANGE: Force Template Selection ---
        # This clears the path every time a relay button is clicked.
        self.template_path.set("")

        self.show_main_interface()

    def show_main_interface(self):
        for widget in self.root.winfo_children():
            widget.destroy()

        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.grid(row=0, column=0, sticky="nsew")
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

        # Header
        title_text = f"{self.relay_config[self.selected_type]['label']} Settings"
        ttk.Label(main_frame, text=title_text, font=('Helvetica', 12, 'bold')).grid(row=0, column=0, columnspan=2,
                                                                                    sticky="w", pady=(0, 15))

        ttk.Button(main_frame, text="← Back", command=self.show_selection_screen).grid(row=0, column=2, sticky="e")

        # File Selectors
        self.create_path_selector(main_frame, "Settings Workbook:", self.xl_path, self.browse_excel, 1)
        self.create_path_selector(main_frame, "RDB Template Dir:", self.template_path, self.browse_template, 2)
        self.create_path_selector(main_frame, "Output Directory:", self.output_path, self.browse_output, 3)

        # Region Selectors (Conditional)
        current_row = 4
        if self.selected_type in self.relay_region_metadata:
            self.build_region_selectors(main_frame, row_idx=4)
            current_row = 5

        # Separator
        ttk.Separator(main_frame, orient='horizontal').grid(row=current_row, column=0, columnspan=3, sticky="ew",
                                                            pady=15)

        # --- NEW CHECKBOX ---
        # This will be on row current_row + 1
        options_frame = ttk.Frame(main_frame)
        options_frame.grid(row=current_row + 1, column=0, columnspan=3, sticky="w")
        ttk.Checkbutton(options_frame, text="Include Comments in RDB", variable=self.include_comments).pack(side="left")

        # Actions
        # The Generate button must now be on row current_row + 2
        generate_btn = ttk.Button(main_frame, text="Generate Settings", command=self.generate_settings,
                                  style='Large.TButton')
        generate_btn.grid(row=current_row + 2, column=0, columnspan=3, pady=10, sticky="ew") # <-- Row index corrected

        ttk.Label(main_frame, textvariable=self.status_var, foreground="blue").grid(row=current_row + 3, column=0,
                                                                                    columnspan=3)

        main_frame.grid_columnconfigure(1, weight=1)

    def create_path_selector(self, parent, label, variable, cmd, row):
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=5)
        entry = ttk.Entry(parent, textvariable=variable)
        entry.grid(row=row, column=1, padx=5, sticky="ew")
        ttk.Button(parent, text="Browse...", command=cmd).grid(row=row, column=2)

    def browse_excel(self):
        filename = filedialog.askopenfilename(title="Select Workbook", filetypes=[("Excel files", "*.xlsx *.xls")])
        if filename: self.xl_path.set(filename)

    def browse_template(self):
        dirname = filedialog.askdirectory(title="Select RDB Template Directory")
        if dirname: self.template_path.set(dirname)

    def browse_output(self):
        dirname = filedialog.askdirectory(title="Select Output Directory")
        if dirname: self.output_path.set(dirname)

    def set_all_regions(self, state):
        """Helper to bulk set region checkboxes"""
        for var in self.region_vars.values():
            var.set(state)

    def build_region_selectors(self, parent, row_idx):
        region_frame = ttk.LabelFrame(parent, text="Region Selection", padding="10")
        region_frame.grid(row=row_idx, column=0, columnspan=3, pady=(10, 0), sticky="ew")

        # --- NEW: Control Buttons (Select All / None) ---
        control_frame = ttk.Frame(region_frame)
        control_frame.grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 10))

        # Link buttons to the set_all_regions helper
        ttk.Button(control_frame, text="Select All", width=10,
                   command=lambda: self.set_all_regions(True)).pack(side="left", padx=(0, 5))
        ttk.Button(control_frame, text="Deselect All", width=10,
                   command=lambda: self.set_all_regions(False)).pack(side="left")

        # --- Region Checkboxes ---
        self.region_vars = {}
        region_meta = self.relay_region_metadata.get(self.selected_type)
        self.current_region_shorthand = region_meta["shorthand"]

        for i, label in enumerate(region_meta["labels"]):
            var = tk.BooleanVar(value=True)
            self.region_vars[label] = var

            # Note: Added '+ 1' to row index to make room for control buttons
            grid_row = (i // 3) + 1
            grid_col = i % 3

            ttk.Checkbutton(region_frame, text=label, variable=var).grid(
                row=grid_row, column=grid_col, sticky="w", padx=10, pady=2
            )

        region_frame.grid_columnconfigure(0, weight=1)
        region_frame.grid_columnconfigure(1, weight=1)
        region_frame.grid_columnconfigure(2, weight=1)

    def generate_settings(self):
        if not all([self.xl_path.get(), self.template_path.get(), self.output_path.get()]):
            messagebox.showerror("Error",
                                 "Please select all required paths.\n(Template must be re-selected for new relay types)")
            return

        try:
            self.status_var.set("Processing...")
            self.root.update()

            excluded_regions = None
            if self.selected_type in self.relay_region_metadata:
                excluded_regions = [
                    self.current_region_shorthand[label]
                    for label, var in self.region_vars.items() if not var.get()
                ]

            gen_settings(
                xl_path=self.xl_path.get(),
                template_path=self.template_path.get(),
                output_path=self.output_path.get(),
                workbook_params=self.workbook_params,
                excluded_regions=excluded_regions,
                include_comments=self.include_comments.get()
            )

            self.status_var.set("Ready")
            messagebox.showinfo("Success", f"Settings generated for {self.selected_type}!")

        except Exception as e:
            self.status_var.set("Error")
            messagebox.showerror("Execution Error", str(e))


if __name__ == "__main__":
    root = tk.Tk()
    app = SettingsGUI(root)
    root.mainloop()