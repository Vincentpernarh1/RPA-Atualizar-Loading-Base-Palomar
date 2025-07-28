import os
import shutil
import pandas as pd
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from openpyxl import load_workbook
from openpyxl.styles import numbers
from openpyxl.formatting.rule import FormulaRule, CellIsRule, ColorScaleRule, DataBarRule, IconSetRule
from openpyxl.styles import PatternFill # Potentially useful for rules that set a fill
from datetime import datetime
from threading import Thread

# Reverse column mapping: English (in loading) → Spanish (in Seguimento)
column_mapping = {
    "Load Number": "LOAD",
    "RO description": "Check flujo",
    "Número de carga externo": "AUTEX N°",
    "Service Level": "Tipo de Servicio",
    "Peso de la carga [kg]": "PESO",
    "Volumen de la carga [m³]": "M3",
    "Medición de la carga [LM]": "SATURACION",
    "Consignor Customer-ID": "COFOR",
    "Consignor Company": "PROVEEDOR",
    "Consignor City": "CIUDAD",
    "Transport Mode": "FLUJO",
    "Means of transport": "TIPO DE VEHICULO COLECTA",
    "Plate Truck": "PLACA COLECTA",
    "Código: Puerto de origen": "CRT",
    "N° de pedido": "NF",
    "Salida del lugar de recogida": "SAÍDA REAL PROVEDOR",
    "Invoice": "FACTURA",
    "TO Comment": "OBSERVACIONES",
    "Latest release date": "LIBERACION EN FRONTERA",
    "Pickup date": "FECHA DE COLECTA iTMS",
    "Delivery date": "FECHA DESCARGA",
    "Llegada al lugar de entrega": "ARRIBO STELLANTIS",
    "ETA": "PREVISION ARRIBO STELLANTIS",
    "ETA al lugar de entrega": "VENTANA ARRIBO STELLANTIS",
    "Service provider Company": "TRANSPORTE"
}

# Spanish columns that are date fields
date_columns = {
    "FECHA DE COLECTA iTMS",
    "FECHA DESCARGA",
    "LIBERACION EN FRONTERA",
    "PREVISION ARRIBO STELLANTIS",
    "VENTANA ARRIBO STELLANTIS",
    "ARRIBO STELLANTIS"
}


class ExcelMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Merger")
        self.frame = tk.Frame(self.root, padx=20, pady=20)
        self.frame.pack()

        tk.Label(self.frame, text="Select folder containing Excel files:").pack()
        tk.Button(self.frame, text="Browse Folder", command=self.browse_folder).pack(pady=10)

        self.progress = ttk.Progressbar(self.frame, orient="horizontal", length=300, mode="determinate")
        self.progress.pack(pady=10)

    def browse_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            Thread(target=self.process_files, args=(folder_selected,)).start()

    def parse_date(self, value):
        """
        Parses a date string and returns a datetime object.
        Returns the original value if parsing fails.
        """
        if pd.isna(value) or value == "":
            return None  # Return None for empty or NaN values

        # Check if the value is already a datetime object (e.g., if Excel already parsed it)
        if isinstance(value, datetime):
            return value

        try:
            # Try parsing with day.month.year format
            return datetime.strptime(str(value).strip(), "%d.%m.%Y")
        except ValueError:
            try:
                # Try parsing with day/month/year format
                return datetime.strptime(str(value).strip(), "%d/%m/%Y")
            except ValueError:
                # If all parsing attempts fail, return the original value
                return value

    def process_files(self, folder_path):
        try:
            self.progress["value"] = 10
            self.root.update_idletasks()

            # Detect files
            loading_file = next((f for f in os.listdir(folder_path) if "loading" in f.lower()), None)
            seguimento_file = next((f for f in os.listdir(folder_path) if "seguimento" in f.lower()), None)

            if not loading_file or not seguimento_file:
                messagebox.showerror("Error", "Required files not found (loading & seguimento).")
                return

            loading_path = os.path.join(folder_path, loading_file)
            seguimento_path = os.path.join(folder_path, seguimento_file)

            # Load loading file
            df_loading = pd.read_excel(loading_path, dtype=str) # Read as string to handle various date formats
            df_loading.rename(columns=column_mapping, inplace=True)

            # Keep only mapped columns
            mapped_cols = [col for col in df_loading.columns if col in column_mapping.values()]
            df_loading = df_loading[mapped_cols]

            if df_loading.empty:
                messagebox.showwarning("Warning", "No valid mapped data found in the loading file.")
                return

            # Parse date fields to datetime objects
            for col in date_columns:
                if col in df_loading.columns:
                    df_loading[col] = df_loading[col].apply(self.parse_date)

            self.progress["value"] = 40
            self.root.update_idletasks()

            # Backup
            backup_path = seguimento_path.replace(".xlsx", "_backup.xlsx")
            shutil.copy(seguimento_path, backup_path)

            # Open Excel and locate sheet
            wb = load_workbook(seguimento_path)
            if "Seguimento" not in wb.sheetnames:
                messagebox.showerror("Error", "'Seguimento' sheet not found.")
                return

            ws = wb["Seguimento"]

            headers_excel = [cell.value for cell in ws[1]]
            col_mapping_excel = {
                col: headers_excel.index(col) + 1
                for col in df_loading.columns if col in headers_excel
            }

            if not col_mapping_excel:
                messagebox.showerror("Error", "None of the mapped columns were found in the sheet.")
                return

            # Get last row with data
            last_data_row = ws.max_row
            for row in range(ws.max_row, 0, -1):
                # Check if the first cell in the row has a value to determine the last data row
                if ws.cell(row=row, column=1).value not in [None, ""]:
                    last_data_row = row
                    break

            start_row = last_data_row + 1

            # Store existing conditional formatting rules and their ranges
            # We will reapply these rules to cover the new rows.
            # Using a list of tuples to store (rule_object, rule_range_string)
            conditional_formatting_rules = []
            for rule_range_str, rules_list in ws.conditional_formatting.items():
                for rule in rules_list:
                    # Store the rule object and its original range string
                    conditional_formatting_rules.append((rule, rule_range_str))

            # Copy formulas from last row
            formula_templates = {}
            for col in range(1, ws.max_column + 1):
                cell = ws.cell(row=last_data_row, column=col)
                if cell.data_type == 'f':
                    formula_templates[col] = cell.value

            self.progress["value"] = 70
            self.root.update_idletasks()

            # Write data
            for i, row in df_loading.iterrows():
                current_row_num = start_row + i
                for col_name, col_index in col_mapping_excel.items():
                    value = row[col_name]
                    cell = ws.cell(row=current_row_num, column=col_index, value=value)
                    
                    # Apply date format if it's a date column and the value is a datetime object
                    if col_name in date_columns and isinstance(value, datetime):
                        cell.number_format = numbers.FORMAT_DATE_DMYSLASH
                    elif col_name in date_columns and value is None:
                        cell.value = "" # Ensure empty cells for None date values

                # Apply formulas to the new row
                for col, formula in formula_templates.items():
                    # Replace the row number in the formula to point to the current new row
                    adjusted_formula = formula.replace(str(last_data_row), str(current_row_num))
                    ws.cell(row=current_row_num, column=col, value=adjusted_formula)

            # Reapply conditional formatting to cover all data, including newly added rows
            # This approach updates the range of existing rules to include the newly added rows.
            max_current_row = ws.max_row
            
            # Clear existing rules to avoid duplicates or outdated ranges
            # It's safer to clear all and re-add with updated ranges.
            # You can also manually delete rules by their ranges if you know them.
            ws.conditional_formatting._cf_rules = {} 

            for rule, original_range_str in conditional_formatting_rules:
                # Parse the original range to get the column parts
                # Example: "A2:Z2" -> "A" and "Z"
                start_col = original_range_str.split(':')[0][0] # Assuming single letter columns
                end_col = original_range_str.split(':')[1][0] # Assuming single letter columns

                # Construct the new range to cover from row 2 (after header) to the new max_current_row
                # Adjust if your conditional formatting doesn't start from row 2
                new_rule_range = f"{start_col}2:{end_col}{max_current_row}"

                # Add the rule back with the updated range
                # The 'rule' object itself needs to be passed, not just its properties.
                ws.conditional_formatting.add(new_rule_range, rule)

            self.progress["value"] = 90
            self.root.update_idletasks()

            # Save
            wb.save(seguimento_path)
            wb.close()

            self.progress["value"] = 100
            messagebox.showinfo("Success", f"{len(df_loading)} rows appended.\nBackup saved as:\n{backup_path}")

        except Exception as e:
            messagebox.showerror("Error", str(e))


# Launch GUI
if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelMergerApp(root)
    root.mainloop()