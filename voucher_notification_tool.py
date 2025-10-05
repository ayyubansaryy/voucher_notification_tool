import os
import time
from io import StringIO
import datetime
import tkinter as tk
import pandas as pd
from tkcalendar import Calendar
from reportlab.lib.pagesizes import A4
import re

def boot_message():
    print("\n══════════════════════════════════════════════════════\nVoucher Notification Text Tool\n══════════════════════════════════════════════════════\nv0.7 (developed by: FEL-89242)\nDefault output location: Desktop\n══════════════════════════════════════════════════════\nPaste data below (with headers) & press 'Enter' twice:\n")

def restart_message():
    os.system("cls" if os.name == "nt" else "clear")
    print("\n🟢 Successfully Restarted\n\nPaste data below (with headers) & press 'Enter' twice:\n══════════════════════════════════════════════════════\n")

def closing():
    print("\n👋 Closing in 5 seconds...")
    time.sleep(3)

#### Format date to string
def format_date(date_input: str) -> str:
    try:
        return datetime.datetime.strptime(date_input, "%d/%m/%Y").strftime("%d %B")
    except ValueError:
        return date_input

#### Pick validity dates from calendar
def pick_date(title="Select a date") -> str:
    root = tk.Tk()
    root.title(title)
    tk.Label(root, text=title, font=("Arial", 16, "bold")).pack(pady=10)
    cal = Calendar(root, selectmode="day", date_pattern="dd/mm/yyyy", font=("Arial", 14))
    cal.pack(pady=10)
    w, h = 500, 400
    x = (root.winfo_screenwidth() - w) // 2
    y = (root.winfo_screenheight() - h) // 2
    root.geometry(f"{w}x{h}+{x}+{y}")
    selected_date = tk.StringVar()
    def confirm():
        selected_date.set(cal.get_date())
        root.destroy()
    tk.Button(root, text="Confirm", font=("Arial", 12, "bold"), padx=20, pady=10,
              command=confirm).pack(pady=10)
    root.lift()
    root.attributes("-topmost", True)
    root.after(100, lambda: root.attributes("-topmost", False))
    root.mainloop()
    return selected_date.get()

#### Get validity dates
def get_validity_dates():
    start = format_date(pick_date("Validity Starts:"))
    end = format_date(pick_date("Validity Ends:"))
    return start, end

#### Build notification text segments
def build_segments(df: pd.DataFrame, start: str, end: str) -> list[str]:
    segments = []

    def format_order_contact(line: str) -> str:
        match = re.match(r"(\w+)\s*(\d+)", line)
        if match:
            order, contact = match.groups()
            return f"{order}\u00A0\u00A0\u00A0{contact}"
        return line

    for serial, (amount, group) in enumerate(df.groupby("Voucher"), start=1):
        code_str = f"SORRY{int(amount)}"
        mov = int(amount) + 49
        order_contact_lines = [
            format_order_contact(f"{row['Order No']} {row['Contact']}")
            for _, row in group.iterrows()
        ]
        lines = [
            f"{serial}. {code_str}",
            *order_contact_lines,
            f"Use coupon {code_str} to get {int(amount)} taka off",
            f"Minimum order: {mov} taka",
            f"Validity: {start} to {end}",
            "Not applicable for Flat discount-providing restaurants\n",
        ]
        segments.append("\n".join(lines))
    return segments

#### Input Data Function
def read_input_data() -> pd.DataFrame:
    while True:
        lines = []
        while True:
            try:
                line = input()
                if line.strip() == "":
                    break
                lines.append(line)
            except EOFError:
                break

        raw = "\n".join(lines).strip()
        if not raw:
            print("🛑 ERROR! No data found!\n")
            try:
                choice = input("🔘 Press ENTER ⏎ to Restart \n\n🔘 Close window to Exit ").strip()
            except EOFError:
                print("\n👋 Window closed. Exiting tool...")
                exit()
            if choice == "":
                restart_message()
                continue
            else:
                restart_message()
                continue

        # Detect header
        first_line = raw.splitlines()[0].strip().lower()
        has_header = any(keyword in first_line for keyword in ["ticket id", "order no", "contact", "voucher"])

        if has_header:
            df = pd.read_csv(StringIO(raw), sep="\t", dtype=str)

        else:
            print("🛑 ERROR! Headers not found!\n")
            try:
                choice = input("🔘 Press ENTER ⏎ to Restart \n\n🔘 Close window to Exit ").strip()
            except EOFError:
                print("\n👋 Window closed. Exiting tool...")
                exit()
            if choice == "":
                restart_message()
                continue
            else:
                restart_message()
                continue

        # Exclude Voucher Given == yes or withdrawn
        if "Voucher Given" in df.columns:
            df = df[~df["Voucher Given"].astype(str).str.strip().str.lower().isin(["yes", "withdrawn"])]

        # Keep only relevant columns
        df = df[[col for col in ["Order No", "Contact", "Voucher"] if col in df.columns]]
        return df

#### Main Program
def main():
    boot_message()
    while True:
        df = read_input_data()
        if df is None:
            
            try:
                choice = input("🔘 Press ENTER ⏎ to Restart \n\n🔘 Close window to Exit ").strip()
            except EOFError:
                print("\n👋 Window closed. Exiting tool...")
                exit()
            if choice == "":
                restart_message()
                continue
            else:
                restart_message()
                continue

        # Check for missing voucher values
        missing_voucher_rows = df[df["Voucher"].isna()]
        if not missing_voucher_rows.empty:
            print("\n🛑 ERROR! Following rows have missing voucher amount:\n")
            for _, row in missing_voucher_rows.iterrows():
                print(f"{row['Order No']}  0{row['Contact']}  {row['Voucher']}")
            
            try:
                choice = input("🔘 Press ENTER ⏎ to Restart \n\n🔘 Close window to Exit ").strip()
            except EOFError:
                print("\n👋 Window closed. Exiting tool...")
                exit()
            if choice == "":
                restart_message()
                continue
            else:
                restart_message()
                continue

        # Check for duplicate contacts
        duplicate_contacts = df[df.duplicated(subset=["Contact"], keep=False)]
        if not duplicate_contacts.empty:
            duplicate_contacts = duplicate_contacts.sort_values(by="Contact")
            print("\nDuplicate contacts found:\n")
            for _, row in duplicate_contacts.iterrows():
                print(f"{row['Order No']} 0{row['Contact']} {int(row['Voucher'])}")
            
            
            try:
                choice = input("\n🔘 Type and enter '𝐂' to continue \n\n🔘 Press ENTER ⏎ to Restart \n\n🔘 Close window to Exit").strip().lower()
            except EOFError:
                print("\n👋 Window closed. Exiting tool...")
                exit()
            if choice == "c":
                print("\n ✌ Allowed duplicates...")
            elif choice == "":
                restart_message()
                continue
            else:
                restart_message()
                continue

        print("\n🔘 Set validity from the pop-up calendar...")
        break

    # Convert the voucher column to integer
    df["Voucher"] = pd.to_numeric(df["Voucher"], errors="coerce").astype("Int64")
    
    # Sort by Voucher
    df = df.sort_values(by="Voucher")

    # Fix contact numbers
    df["Contact"] = df["Contact"].apply(lambda x: x if len(x) != 10 else "0" + x)

    # Get validity dates
    start_date, end_date = get_validity_dates()

    # Build segments
    segments = build_segments(df, start_date, end_date)
    final_text = f"Need to send notification for the coupon list below:\n\n{'\n'.join(segments)}"

    # Export as Notepad text file
    output_folder = os.path.join(os.path.expanduser("~"), "Desktop")
    file_name = f"Coupons_notification_{start_date.replace(' ', '_')}_to_{end_date.replace(' ', '_')}.txt"
    output_path = os.path.join(output_folder, file_name)
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(final_text)

    print("\n:::::: ✨  Notepad text file generated  ╰┈➤  📁", output_path)
    time.sleep(2)

    while True:
        try:
            choice = input("\n\n🔘 Press ENTER ⏎ to Restart \n\n🔘 Close window to Exit ").strip()
        except EOFError:
            print("\n👋 Window closed. Exiting tool...")
            exit()
        if choice == "":
            os.system("cls" if os.name == "nt" else "clear")
            print("\n🟢 Successfully Restarted")
            return main()
        else:
            os.system("cls" if os.name == "nt" else "clear")
            print("\n🟢 Successfully Restarted")
            return main()

if __name__ == "__main__":
    main()
