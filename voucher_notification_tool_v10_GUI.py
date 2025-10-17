# Compilation command: pyinstaller --onefile --windowed --distpath . "voucher_notification_tool_v10_GUI.py"

import os
import re
import tkinter as tk
from io import StringIO
import datetime
import pandas as pd
from tabulate import tabulate
from tkcalendar import Calendar
import customtkinter as ctk

# --- Core Logic from Original Script (Slightly Modified for GUI) ---
# These functions are the "engine" of your application.

def format_date(date_input: str) -> str:
    """Formats a dd/mm/yyyy string to 'dd Month' format."""
    try:
        return datetime.datetime.strptime(date_input, "%d/%m/%Y").strftime("%d %B")
    except (ValueError, TypeError):
        return None

def build_segments(df: pd.DataFrame, start: str, end: str) -> list[str]:
    """Builds the final notification text segments based on processed data."""
    segments = []

    def format_order_contact(line: str) -> str:
        match = re.match(r"(\w+)\s*(\d+)", line)
        if match:
            order, contact = match.groups()
            return f"{order}\u00A0\u00A0\u00A0{contact}"
        return line

    start_day, start_month = start.split()
    end_day, end_month = end.split()

    def get_day_with_suffix(d):
        day_num = int(d)
        if 11 <= day_num <= 13:
            return f"{d}th"
        elif day_num % 10 == 1:
            return f"{d}st"
        elif day_num % 10 == 2:
            return f"{d}nd"
        elif day_num % 10 == 3:
            return f"{d}rd"
        else:
            return f"{d}th"

    start_date_str = f"{get_day_with_suffix(start_day)} {start_month}"
    end_date_str = f"{get_day_with_suffix(end_day)} {end_month}"

    for serial, (amount, group) in enumerate(df.groupby("Voucher"), start=1):
        code_str = f"SORRY{int(amount)}"
        mov = int(amount) + 49
        order_contact_lines = [
            format_order_contact(f"{row['Order No']} {row['Contact']}")
            for _, row in group.iterrows()
        ]
        lines = [
            f"{serial}. {code_str}\n",
            *order_contact_lines,
            f"\nUse coupon {code_str} to get {int(amount)} taka off",
            f"Minimum order: {mov} taka",
            f"Validity: {start_date_str} to {end_date_str}",
            "Not applicable for Flat discount-providing restaurants\n",
        ]
        segments.append("\n".join(lines))
    return segments


# --- The Modern GUI Application Class ---

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        # --- Font Sizes (Increased for better visibility) ---
        self.TITLE_FONT_SIZE = 26
        self.HEADER_FONT_SIZE = 16
        self.BUTTON_FONT_SIZE = 16
        self.TEXTBOX_FONT_SIZE = 14
        self.STATUS_FONT_SIZE = 14

        # --- Window Configuration ---
        self.title("Voucher Notification Tool v0.9 (GUI)")
        self.geometry("750x750")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        # --- State Variables ---
        self.start_date = None
        self.end_date = None

        # --- Store the dataframe after previewing
        self.processed_df = None 

        # --- Widgets ---
        self.title_label = ctk.CTkLabel(self, text="Voucher Notification Tool", font=ctk.CTkFont(size=self.TITLE_FONT_SIZE, weight="bold"))
        self.title_label.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="ew")
        
        # --- Tab View for Step-by-Step process ---
        self.tab_view = ctk.CTkTabview(self, anchor="w")
        self.tab_view.grid(row=1, column=0, padx=20, pady=20, sticky="nsew")
        
        self.input_tab = self.tab_view.add("Step 1: Input Data")
        self.preview_tab = self.tab_view.add("Step 2: Preview & Generate")
        
        # --- Configure Grid for Tabs ---
        self.input_tab.grid_columnconfigure(0, weight=1)
        self.input_tab.grid_rowconfigure(2, weight=1)
        self.preview_tab.grid_columnconfigure(0, weight=1)
        self.preview_tab.grid_rowconfigure(0, weight=1)

        # --- Widgets for Input Tab ---
        self.controls_frame = ctk.CTkFrame(self.input_tab)
        self.controls_frame.grid(row=0, column=0, padx=0, pady=0, sticky="ew")
        self.controls_frame.grid_columnconfigure((0, 1), weight=1)

        self.session_label = ctk.CTkLabel(self.controls_frame, text="Voucher Session", font=ctk.CTkFont(size=self.HEADER_FONT_SIZE, weight="bold"))
        self.session_label.grid(row=0, column=0, padx=20, pady=(10, 5), sticky="w")
        self.session_var = ctk.StringVar(value="Evening")
        self.session_selector = ctk.CTkSegmentedButton(
            self.controls_frame, values=["Morning", "Evening"], variable=self.session_var,
            font=ctk.CTkFont(size=self.BUTTON_FONT_SIZE)
        )
        self.session_selector.grid(row=1, column=0, padx=20, pady=(0, 20), sticky="ew")

        self.date_label = ctk.CTkLabel(self.controls_frame, text="Voucher Validity", font=ctk.CTkFont(size=self.HEADER_FONT_SIZE, weight="bold"))
        self.date_label.grid(row=0, column=1, padx=20, pady=(10, 5), sticky="w")
        self.date_buttons_frame = ctk.CTkFrame(self.controls_frame, fg_color="transparent")
        self.date_buttons_frame.grid(row=1, column=1, padx=20, pady=(0, 20), sticky="ew")
        self.date_buttons_frame.grid_columnconfigure((0, 1), weight=1)
        
        self.start_date_button = ctk.CTkButton(self.date_buttons_frame, text="Select start date", command=self.pick_start_date, font=ctk.CTkFont(size=self.BUTTON_FONT_SIZE))
        self.start_date_button.grid(row=0, column=0, padx=(0, 5), sticky="ew")
        
        self.end_date_button = ctk.CTkButton(self.date_buttons_frame, text="Select end date", command=self.pick_end_date, font=ctk.CTkFont(size=self.BUTTON_FONT_SIZE))
        self.end_date_button.grid(row=0, column=1, padx=(5, 0), sticky="ew")

        # --- Input Label and Clear Button Frame ---
        self.input_header_frame = ctk.CTkFrame(self.input_tab, fg_color="transparent")
        self.input_header_frame.grid(row=1, column=0, padx=0, pady=(10, 5), sticky="ew")
        self.input_header_frame.grid_columnconfigure(0, weight=1)
        self.input_label = ctk.CTkLabel(self.input_header_frame, text="Paste Data Below (with headers)", font=ctk.CTkFont(size=self.HEADER_FONT_SIZE, weight="bold"))
        self.input_label.grid(row=0, column=0, sticky="w")
        self.clear_button = ctk.CTkButton(self.input_header_frame, text="Clear Data", width=100, command=self.clear_all_data, font=ctk.CTkFont(size=self.BUTTON_FONT_SIZE))
        self.clear_button.grid(row=0, column=1, sticky="e")

        self.input_textbox = ctk.CTkTextbox(self.input_tab, height=300, font=("Consolas", self.TEXTBOX_FONT_SIZE))
        self.input_textbox.grid(row=2, column=0, padx=0, pady=5, sticky="nsew")
        
        self.preview_button = ctk.CTkButton(self.input_tab, text="Preview Data ➡️", font=ctk.CTkFont(size=self.HEADER_FONT_SIZE, weight="bold"), height=40, command=self.show_preview)
        self.preview_button.grid(row=3, column=0, padx=0, pady=20, sticky="ew")

        # --- Widgets for Preview Tab ---
        self.preview_textbox = ctk.CTkTextbox(self.preview_tab, font=("Consolas", self.TEXTBOX_FONT_SIZE), state="disabled")
        self.preview_textbox.grid(row=0, column=0, padx=0, pady=10, sticky="nsew")

        self.generate_button = ctk.CTkButton(self.preview_tab, text="Generate Notification File ➡️", font=ctk.CTkFont(size=self.HEADER_FONT_SIZE, weight="bold"), height=40, command=self.generate_file)
        self.generate_button.grid(row=1, column=0, padx=0, pady=(10, 18), sticky="ew")
        
        # --- Status Bar ---
        self.status_label = ctk.CTkLabel(self, text="Step 1: Please select validity dates and paste data (with headers).", text_color="gray60", font=ctk.CTkFont(size=self.STATUS_FONT_SIZE)) # Increased
        self.status_label.grid(row=2, column=0, padx=20, pady=(0, 10), sticky="w")

    def pick_date_dialog(self, title):
        dialog = ctk.CTkToplevel(self)
        dialog.title(title)
        dialog.geometry("350x350")
        dialog.transient(self)
        dialog.grab_set()
        
        selected_date = None
        def on_date_select(date):
            nonlocal selected_date
            selected_date = date
            dialog.destroy()
            
        cal = Calendar(dialog, selectmode="day", date_pattern="dd/mm/yyyy", font=("Segoe UI", 14))
        cal.pack(pady=0, padx=0, fill="both", expand=True)
        confirm_button = ctk.CTkButton(dialog, text="OK", command=lambda: on_date_select(cal.get_date()), font=ctk.CTkFont(size=self.BUTTON_FONT_SIZE))
        confirm_button.pack(pady=10)
        self.wait_window(dialog)
        return selected_date

    def pick_start_date(self):
        date_str = self.pick_date_dialog("Select Validity Start Date")
        if date_str:
            self.start_date = format_date(date_str)
            self.start_date_button.configure(text=f"Start: {self.start_date}")
            self.update_status("Start date selected.", "gray60")

    def pick_end_date(self):
        date_str = self.pick_date_dialog("Select Validity End Date")
        if date_str:
            self.end_date = format_date(date_str)
            self.end_date_button.configure(text=f"End: {self.end_date}")
            self.update_status("End date selected.", "gray60")

    def update_status(self, message, color):
        self.status_label.configure(text=message, text_color=color)

    def clear_all_data(self):
        """Clears all input fields and resets the application state."""
        self.input_textbox.delete("1.0", "end")
        
        self.preview_textbox.configure(state="normal")
        self.preview_textbox.delete("1.0", "end")
        self.preview_textbox.configure(state="disabled")
        
        self.processed_df = None
        self.update_status("Inputs cleared. Ready to paste new data.", "gray60")
        self.tab_view.set("Step 1: Input Data")

    def show_preview(self):
        self.processed_df = None

        if not self.start_date or not self.end_date:
            self.update_status("Error: Please select both a start and end date.", "yellow")
            return

        raw_data = self.input_textbox.get("1.0", "end-1c").strip()
        if not raw_data:
            self.update_status("Error: Input data cannot be empty.", "yellow")
            return

        try:
            first_line = raw_data.splitlines()[0].strip().lower()
            if not any(k in first_line for k in ["order no", "contact", "voucher"]):
                self.update_status("Error: Headers not found in the first line.", "yellow")
                return

            # Parse TSV-like text
            df = pd.read_csv(StringIO(raw_data), sep="\t", dtype=str)

            # Remove withdrawn or already given vouchers
            if "Voucher Given" in df.columns:
                df = df[~df["Voucher Given"].astype(str).str.strip().str.lower().isin(["yes", "withdrawn"])]

            # Prepare and clean columns
            expected_cols = ["Order No", "Contact", "Voucher"]
            df = df[[c for c in expected_cols if c in df.columns]]

            # Convert numeric
            df["Voucher"] = pd.to_numeric(df["Voucher"], errors="coerce")

            # Identify invalids
            invalid_voucher_mask = df["Voucher"].isnull()
            missing_order_mask = df["Order No"].isnull() | (df["Order No"].astype(str).str.strip() == "")
            duplicate_mask = df.duplicated(subset=["Contact"], keep=False)

            invalid_voucher_count = invalid_voucher_mask.sum()
            missing_order_count = missing_order_mask.sum()
            duplicate_count = duplicate_mask.sum()

            # --- Build Invalid DataFrames for Each Error Type ---
            invalid_groups = {}

            if invalid_voucher_count > 0:
                invalid_voucher_df = df.loc[invalid_voucher_mask].copy()
                invalid_voucher_df["Reason"] = "Invalid Voucher"
                invalid_groups["Voucher Invalid"] = invalid_voucher_df

            if missing_order_count > 0:
                missing_order_df = df.loc[missing_order_mask].copy()
                missing_order_df["Reason"] = "Missing Order ID"
                invalid_groups["Order ID Missing"] = missing_order_df

            if duplicate_count > 0:
                duplicate_df = df.loc[duplicate_mask].copy()
                duplicate_df["Reason"] = "Duplicate Contact"
                invalid_groups["Duplicate Contact"] = duplicate_df

            # Combine all invalids into one if you still need the full set
            invalid_df = pd.concat(invalid_groups.values(), ignore_index=True) if invalid_groups else pd.DataFrame()

            # Filter valid entries
            valid_df = df[~invalid_voucher_mask & ~missing_order_mask].copy()

            if valid_df.empty:
                self.update_status("Error: No valid entries found.", "yellow")
                return

            summary_parts = []

            summary_data = [
                ["Total Rows", len(df)],
                ["Valid Entries", len(valid_df)],
                ["Invalid Vouchers", invalid_voucher_count],
                ["Missing Order IDs", missing_order_count],
                ["Duplicate Contacts", duplicate_count],
            ]
            summary_table = tabulate(
                summary_data,
                tablefmt="fancy_grid",
                showindex=False,
            )
            summary_parts.append(f"**Data Summary**:\n{summary_table}\n\n")

            # --- Invalid Data Preview (NEW) ---
            if not invalid_df.empty:
                invalid_table = tabulate(invalid_df, headers="keys", tablefmt="fancy_grid", showindex=False)
                summary_parts.append(f"⚠️ Invalid Data Preview:\n{invalid_table}\n\n")

            # --- Voucher Distribution ---
            if "Voucher" in valid_df.columns and not valid_df["Voucher"].isnull().all():
                voucher_counts = valid_df["Voucher"].dropna().astype(int).value_counts().sort_index()
                voucher_data = [[voucher, count] for voucher, count in voucher_counts.items()]
                voucher_summary_table = tabulate(
                    voucher_data,
                    headers=["Voucher", "Count"],
                    tablefmt="rounded_outline",
                    showindex=False,
                )
                summary_parts.append(f"Voucher Distribution:\n{voucher_summary_table}\n\n")

            # --- Valid Data Preview ---
            raw_data_table = tabulate(valid_df, headers="keys", tablefmt="rounded_outline", showindex=False)
            summary_parts.append(f"✅ Valid Data Preview:\n{raw_data_table}\n")            

            preview_content = "\n".join(summary_parts)

            self.preview_textbox.configure(state="normal")
            self.preview_textbox.delete("1.0", "end")
            self.preview_textbox.insert("1.0", preview_content)
            self.preview_textbox.configure(state="disabled")

            self.processed_df = valid_df  # Store only valid rows
            self.tab_view.set("Step 2: Preview & Generate")
            self.update_status("✅ Preview generated with validation summary.", "white")

        except Exception as e:
            self.update_status(f"Error parsing data: {e}", "yellow")



    def generate_file(self):
        if self.processed_df is None or self.processed_df.empty:
            self.update_status("Error: No valid data to process. Please go back to Step 1.", "yellow")
            return

        df = self.processed_df.copy()

        if df["Voucher"].isnull().any():
            self.update_status("Error: One or more rows have a missing/invalid 'Voucher' amount.", "yellow")
            return
        if df["Order No"].isnull().any() or (df["Order No"] == "").any():
            self.update_status("Error: One or more rows have a missing 'Order No'.", "yellow")
            return
        if df.duplicated(subset=["Contact"]).any():
            self.update_status("! Warning: Duplicate contacts found. Processing anyway.", "orange")
        
        try:
            df["Contact"] = df["Contact"].apply(lambda x: str(x) if len(str(x)) != 10 else "0" + str(x))
            df["Voucher"] = df["Voucher"].astype(int)
            df = df.sort_values(by="Voucher")
            
            segments = build_segments(df, self.start_date, self.end_date)
            
            output_folder = os.path.join(os.path.expanduser("~"), "Desktop")
            user_session = self.session_var.get()
            file_name = f"{user_session}_{self.start_date.replace(' ', '_')}_to_{self.end_date.replace(' ', '_')}.txt"
            output_path = os.path.join(output_folder, file_name)

            with open(output_path, "w", encoding="utf-8") as f:
                f.write(f"Need to send notification for the coupon list below:\n\n{'\n\n'.join(segments)}")

            self.update_status(f"✨ Notepad text file generated  ╰┈➤  📁 {output_path}", "white")
            
            if os.name == 'nt':
                os.startfile(output_path)

        except Exception as e:
            self.update_status(f"Error generating file: {e}", "yellow")

if __name__ == "__main__":
    app = App()
    app.mainloop()